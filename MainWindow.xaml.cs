using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Windowing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EXOKit.Services;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace EXOKit
{
    /// <summary>
    /// Main window for EXOKit. Wires the ported toolkit services (EXO/Graph connection, mailbox and
    /// group permission workflows, Bookings access, recipient lookup) to a WinUI TabView shell that
    /// mirrors the original WinForms tool's tabs, enforcing the same EXO-then-Graph connection order
    /// and streaming Logger output into the log panel.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private readonly ToolConfig _config;
        private readonly ExoPowerShellService _exo = new();
        private AuthService _authService;
        private GraphService _graphService;
        private readonly SnapshotService _snapshotService = new();
        private SnapshotRestoreService _snapshotRestoreService;
        private readonly ObservableCollection<SnapshotListItem> _snapshotItems = new();
        private readonly MailboxPermissionService _mailboxPermissionService;
        private GroupMembershipService _groupMembershipService;
        private BookingsService _bookingsService;
        private readonly RecipientLookupService _recipientLookupService;
        private readonly ReportingService _reportingService;
        private readonly SharedMailboxService _sharedMailboxService;
        private GroupCreationService _groupCreationService;
        private readonly GroupSettingsService _groupSettingsService;
        private ServiceNowService? _serviceNowService;
        private readonly ObservableCollection<ReportRow> _reportResults = new();
        private double _lastExpandedOutputHeight = 220;
        private bool _outputCollapsed;
        private bool _showWarnings = true;
        private string? _loadedGroupSettingsIdentity;
        private bool _operationInProgress;
        private System.Threading.CancellationTokenSource? _operationCancellation;
        private readonly List<Control> _disabledOperationControls = new();

        public MainWindow()
        {
            InitializeComponent();

            AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "toolbox.ico"));
            ApplyDarkTitleBar();

            string? configError = null;
            try { _config = ConfigService.Load(); }
            catch (Exception exception)
            {
                _config = ConfigService.CreateDefault();
                configError = exception.Message;
            }
            _snapshotService.TenantIdProvider = () => _exo.ConnectedTenantId;
            _authService = new AuthService(_config.Settings.GraphApi.Scopes.ToArray(), GetWindowHandle, _config.Settings.GraphApi.ClientId, _config.Settings.GraphApi.TenantId);
            _graphService = new GraphService(_authService, _config.Settings.GraphApi.Scopes.ToArray());
            _mailboxPermissionService = new MailboxPermissionService(_exo, _snapshotService);
            _groupMembershipService = new GroupMembershipService(_exo, _graphService, _snapshotService);
            _snapshotRestoreService = new SnapshotRestoreService(_exo, _graphService);
            _bookingsService = new BookingsService(_exo, _graphService, _config);
            _recipientLookupService = new RecipientLookupService(_exo);
            _reportingService = new ReportingService(_exo);
            _sharedMailboxService = new SharedMailboxService(_exo);
            _groupCreationService = new GroupCreationService(_exo, _graphService);
            _groupSettingsService = new GroupSettingsService(_exo, _snapshotService);
            _serviceNowService = _config.Settings.ServiceNow != null ? new ServiceNowService(_config.Settings.ServiceNow, _config.Settings.GraphApi) : null;

            ListViewReportResults.ItemsSource = _reportResults;
            ListViewSnapshots.ItemsSource = _snapshotItems;

            LoadSettingsIntoUi();

            Logger.LogEntryWritten += OnLogEntryWritten;
            TextBoxGroupSettingsIdentity.TextChanged += (_, _) => _loadedGroupSettingsIdentity = null;
            if (configError != null)
            {
                Logger.Log($"Configuration could not be loaded: {configError}. Correct and save Settings; the original file is preserved as config.json.bak on save.", LogType.Error);
                MainNavigationView.SelectedItem = NavItemSettings;
            }
            Closed += async (_, _) =>
            {
                Logger.LogEntryWritten -= OnLogEntryWritten;
                try { await _exo.ShutdownAsync(); }
                catch (Exception exception) { Logger.Log($"Session shutdown: {exception.Message}", LogType.Warning); }
            };

            UpdateConnectionStatus();
            var version = GetAppVersion();
            Title = $"EXOKit v{version}";
            TextBlockVersion.Text = $"Version {version}";
            AppWindow.Closing += (_, args) => args.Cancel = _operationInProgress;
        }

        private static string GetAppVersion()
        {
            try
            {
                var version = Windows.ApplicationModel.Package.Current.Id.Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
            catch (Exception exception) when (exception is InvalidOperationException or COMException)
            {
                return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "development";
            }
        }

        private async Task RunUiOperationAsync(Func<Task> operation)
        {
            if (_operationInProgress) return;
            _operationInProgress = true;
            _operationCancellation = new System.Threading.CancellationTokenSource();
            _exo.OperationCancellationToken = _operationCancellation.Token;
            _authService.OperationCancellationToken = _operationCancellation.Token;
            ButtonCancelOperation.Visibility = Visibility.Visible;
            SetOperationControlsEnabled(false);
            try { await operation(); }
            catch (OperationCanceledException)
            {
                Logger.Log("Operation cancelled. Applied changes are not rolled back; review results and recovery snapshots before retrying.", LogType.Warning);
            }
            catch (Exception exception)
            {
                Logger.Log($"Operation stopped: {exception.Message}. Review applied changes before retrying.", LogType.Error);
                await ShowMessageAsync(exception.Message, "Operation Failed");
            }
            finally
            {
                _operationInProgress = false;
                _exo.OperationCancellationToken = default;
                _authService.OperationCancellationToken = default;
                _operationCancellation.Dispose();
                _operationCancellation = null;
                ButtonCancelOperation.Visibility = Visibility.Collapsed;
                SetOperationControlsEnabled(true);
                UpdateConnectionStatus();
            }
        }

        private async void ButtonCheckUpdates_Click(object sender, RoutedEventArgs args) =>
            await RunUiOperationAsync(async () =>
            {
                var feed = _config.Settings.UpdateFeedUrl;
                if (!Uri.TryCreate(feed, UriKind.Absolute, out var uri) || uri.Scheme != "https" || !string.IsNullOrEmpty(uri.UserInfo))
                    throw new InvalidOperationException("Configure a trusted HTTPS .appinstaller feed in Settings first.");
                if (!await ShowConfirmAsync($"Open the update feed at {uri.Host}? Windows App Installer will validate the package signature.", "Check for Updates")) return;
                await Windows.System.Launcher.LaunchUriAsync(uri);
            });

        private IntPtr GetWindowHandle() => WindowNative.GetWindowHandle(this);

        private void ButtonCancelOperation_Click(object sender, RoutedEventArgs args)
        {
            _operationCancellation?.Cancel();
            Logger.Log("Cancellation requested. Waiting for the active service call to finish; applied changes will not be rolled back.", LogType.Warning);
        }

        private void SetOperationControlsEnabled(bool enabled)
        {
            if (enabled)
            {
                foreach (var control in _disabledOperationControls) control.IsEnabled = true;
                _disabledOperationControls.Clear();
                return;
            }

            void DisableControls(DependencyObject parent)
            {
                if (parent is Control control)
                {
                    if (control.IsEnabled)
                    {
                        _disabledOperationControls.Add(control);
                        control.IsEnabled = false;
                    }
                    return;
                }
                for (var index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
                    DisableControls(VisualTreeHelper.GetChild(parent, index));
            }

            DisableControls(OperationPanels);
            DisableControls(ConnectionControls);
            DisableControls(TextBoxServiceNowTicketNumber);
            foreach (var item in MainNavigationView.MenuItems.OfType<NavigationViewItem>()) DisableControls(item);
        }

        private static readonly Regex DeviceCodeLineRegex = new(
            @"^(?<prefix>\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]\s*)?To sign in, use a web browser to open the page (?<url>\S+) and enter the code (?<code>\S+) to authenticate\.$",
            RegexOptions.Compiled);

        private void OnLogEntryWritten(string line, LogType type)
        {
            if (type == LogType.Warning && !_showWarnings)
            {
                return;
            }

            DispatcherQueue.TryEnqueue(() =>
            {
                AppendLogLine(line, type);
                LogScrollViewer.UpdateLayout();
                LogScrollViewer.ChangeView(null, LogScrollViewer.ScrollableHeight, null, true);
            });
        }

        /// <summary>
        /// Toggles whether warning-level log lines (e.g. benign Exchange Online replication warnings) are
        /// shown in the log panel. Warnings are suppressed by default to reduce noise; checking the box
        /// re-renders the full log history including any previously hidden warnings.
        /// </summary>
        private void CheckBoxShowWarnings_Changed(object sender, RoutedEventArgs e)
        {
            if (CheckBoxShowWarnings == null || TextBlockLog == null || LogScrollViewer == null) return;
            _showWarnings = CheckBoxShowWarnings.IsChecked == true;

            TextBlockLog.Inlines.Clear();
            foreach (var (historyLine, historyType) in Logger.GetHistory())
            {
                if (historyType == LogType.Warning && !_showWarnings)
                {
                    continue;
                }

                AppendLogLine(historyLine, historyType);
            }

            LogScrollViewer.UpdateLayout();
            LogScrollViewer.ChangeView(null, LogScrollViewer.ScrollableHeight, null, true);
        }

        private void AppendLogLine(string line, LogType type)
        {
            var match = DeviceCodeLineRegex.Match(line);
            if (!match.Success)
            {
                TextBlockLog.Inlines.Add(new Run
                {
                    Text = line + Environment.NewLine,
                    Foreground = new SolidColorBrush(Logger.GetColorForType(type))
                });
            }
            else
            {
                if (match.Groups["prefix"].Success)
                {
                    TextBlockLog.Inlines.Add(new Run { Text = match.Groups["prefix"].Value });
                }

                TextBlockLog.Inlines.Add(new Run { Text = "To sign in, use a web browser to open the page " });

                var url = match.Groups["url"].Value;
                var hyperlink = new Hyperlink
                {
                    NavigateUri = new Uri(url),
                    Foreground = new SolidColorBrush(Microsoft.UI.Colors.LimeGreen)
                };
                hyperlink.Inlines.Add(new Run { Text = url });
                TextBlockLog.Inlines.Add(hyperlink);

                TextBlockLog.Inlines.Add(new Run { Text = " and enter the code " });

                TextBlockLog.Inlines.Add(new Run
                {
                    Text = match.Groups["code"].Value,
                    Foreground = new SolidColorBrush(Microsoft.UI.Colors.Yellow),
                    FontWeight = Microsoft.UI.Text.FontWeights.Bold
                });

                TextBlockLog.Inlines.Add(new Run { Text = " to authenticate." + Environment.NewLine });

                TryCopyDeviceCodeToClipboard(match.Groups["code"].Value);
            }

            if (type == LogType.Ticket)
            {
                TextBlockTicketNotes.Inlines.Add(new Run
                {
                    Text = line + Environment.NewLine,
                    Foreground = new SolidColorBrush(Logger.GetColorForType(type))
                });
                TicketNotesScrollViewer.UpdateLayout();
                TicketNotesScrollViewer.ChangeView(null, TicketNotesScrollViewer.ScrollableHeight, null, true);

            }
        }

        private string? _lastCopiedDeviceCode;

        /// <summary>
        /// Automatically copies a freshly detected device sign-in code to the clipboard, so the user can
        /// immediately paste it into the browser tab opened via the "To sign in..." link without having
        /// to manually select/copy the highlighted code text. Guards against re-copying the same code if
        /// the log line is somehow processed more than once.
        /// </summary>
        private void TryCopyDeviceCodeToClipboard(string code)
        {
            if (string.IsNullOrWhiteSpace(code) || string.Equals(code, _lastCopiedDeviceCode, StringComparison.Ordinal))
            {
                return;
            }

            try
            {
                var dataPackage = new DataPackage();
                dataPackage.SetText(code);
                Clipboard.SetContent(dataPackage);
                _lastCopiedDeviceCode = code;
                Logger.Log($"Device code '{code}' copied to clipboard.", LogType.Success);
            }
            catch (Exception ex)
            {
                Logger.Log($"Could not copy device code to clipboard: {ex.Message}", LogType.Warning);
            }
        }

        // --- Title bar ---

        private void ApplyDarkTitleBar()
        {
            if (AppWindowTitleBar.IsCustomizationSupported())
            {
                var titleBar = AppWindow.TitleBar;
                titleBar.BackgroundColor = Windows.UI.Color.FromArgb(255, 32, 32, 32);
                titleBar.InactiveBackgroundColor = Windows.UI.Color.FromArgb(255, 32, 32, 32);
                titleBar.ForegroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255);
                titleBar.InactiveForegroundColor = Windows.UI.Color.FromArgb(255, 200, 200, 200);
                titleBar.ButtonBackgroundColor = Windows.UI.Color.FromArgb(255, 32, 32, 32);
                titleBar.ButtonInactiveBackgroundColor = Windows.UI.Color.FromArgb(255, 32, 32, 32);
                titleBar.ButtonForegroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255);
                titleBar.ButtonInactiveForegroundColor = Windows.UI.Color.FromArgb(255, 200, 200, 200);
                titleBar.ButtonHoverBackgroundColor = Windows.UI.Color.FromArgb(255, 60, 60, 60);
                titleBar.ButtonHoverForegroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255);
                titleBar.ButtonPressedBackgroundColor = Windows.UI.Color.FromArgb(255, 80, 80, 80);
                titleBar.ButtonPressedForegroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255);
            }
        }

        // --- Output collapse/expand ---

        private void ButtonToggleOutputCollapse_Click(object sender, RoutedEventArgs e)
        {
            if (_outputCollapsed)
            {
                OutputRow.Height = new GridLength(_lastExpandedOutputHeight);
                OutputPivot.Visibility = Visibility.Visible;
                IconToggleOutputCollapse.Glyph = "\uE70D";
                _outputCollapsed = false;
            }
            else
            {
                _lastExpandedOutputHeight = OutputRow.ActualHeight > 0 ? OutputRow.ActualHeight : _lastExpandedOutputHeight;
                OutputPivot.Visibility = Visibility.Collapsed;
                OutputRow.Height = GridLength.Auto;
                IconToggleOutputCollapse.Glyph = "\uE70E";
                _outputCollapsed = true;
            }
        }

        // --- ServiceNow auto-close ---

        private async Task TryAutoCloseServiceNowTicketAsync()
        {
            _operationCancellation?.Token.ThrowIfCancellationRequested();
            var ticketNumber = TextBoxServiceNowTicketNumber.Text.Trim();
            if (string.IsNullOrWhiteSpace(ticketNumber) || _serviceNowService == null)
            {
                return;
            }

            var notes = Logger.GetLastTicketNotes();
            if (notes.Count == 0)
            {
                return;
            }

            var notesText = string.Join(Environment.NewLine, notes);
            if (!await ShowConfirmAsync($"Close ticket '{ticketNumber}' with these verified operation notes?\n\n{notesText}", "Confirm Ticket Closure")) return;
            Logger.Log($"[ServiceNow] Attempting to update ticket '{ticketNumber}'...", LogType.Info);
            try
            {
                var result = await _serviceNowService.CloseTaskAsync(ticketNumber, notesText, notesText);
                if (result.Success)
                {
                    Logger.Log($"[ServiceNow] {result.Message}", LogType.Success);
                    TextBoxServiceNowTicketNumber.Text = string.Empty;
                }
                else
                {
                    Logger.Log($"[ServiceNow] WARNING: {result.Message}", LogType.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[ServiceNow] ERROR: {ex.Message}", LogType.Error);
            }
        }

        // --- Navigation ---

        private void SetNavIcons(bool show)
        {
            NavItemMailboxes.Icon = show ? new FontIcon { Glyph = "M", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemCreateMailbox.Icon = show ? new FontIcon { Glyph = "C", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemResources.Icon = show ? new FontIcon { Glyph = "R", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemGroups.Icon = show ? new FontIcon { Glyph = "G", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemCreateGroup.Icon = show ? new FontIcon { Glyph = "CG", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemGroupSettings.Icon = show ? new FontIcon { Glyph = "GS", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemBookings.Icon = show ? new FontIcon { Glyph = "B", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemRecipientLookup.Icon = show ? new FontIcon { Glyph = "RL", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemReporting.Icon = show ? new FontIcon { Glyph = "Rp", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemSnapshots.Icon = show ? new FontIcon { Glyph = "Sh", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
            NavItemSettings.Icon = show ? new FontIcon { Glyph = "S", FontFamily = new FontFamily("Segoe UI Semibold") } : null;
        }

        private void MainNavigationView_PaneOpening(NavigationView sender, object args)
        {
            SetNavIcons(false);
        }

        private void MainNavigationView_PaneClosing(NavigationView sender, NavigationViewPaneClosingEventArgs args)
        {
            SetNavIcons(true);
        }

        private void MainNavigationView_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
        {
            var tag = (args.SelectedItem as NavigationViewItem)?.Tag as string;

            PanelMailboxes.Visibility = Visibility.Collapsed;
            PanelCreateMailbox.Visibility = Visibility.Collapsed;
            PanelResources.Visibility = Visibility.Collapsed;
            PanelGroups.Visibility = Visibility.Collapsed;
            PanelCreateGroup.Visibility = Visibility.Collapsed;
            PanelGroupSettings.Visibility = Visibility.Collapsed;
            PanelBookings.Visibility = Visibility.Collapsed;
            PanelRecipientLookup.Visibility = Visibility.Collapsed;
            PanelReporting.Visibility = Visibility.Collapsed;
            PanelSnapshots.Visibility = Visibility.Collapsed;
            PanelSettings.Visibility = Visibility.Collapsed;

            switch (tag)
            {
                case "CreateMailbox":
                    PanelCreateMailbox.Visibility = Visibility.Visible;
                    if (ComboBoxNewMbDomain.ItemsSource == null)
                    {
                        _ = RunUiOperationAsync(RefreshAcceptedDomainsAsync);
                    }
                    break;
                case "Resources":
                    PanelResources.Visibility = Visibility.Visible;
                    break;
                case "Groups":
                    PanelGroups.Visibility = Visibility.Visible;
                    break;
                case "CreateGroup":
                    PanelCreateGroup.Visibility = Visibility.Visible;
                    if (ComboBoxNewGroupDomain.ItemsSource == null)
                    {
                        _ = RunUiOperationAsync(RefreshAcceptedDomainsAsync);
                    }
                    break;
                case "GroupSettings":
                    PanelGroupSettings.Visibility = Visibility.Visible;
                    break;
                case "Bookings":
                    PanelBookings.Visibility = Visibility.Visible;
                    break;
                case "RecipientLookup":
                    PanelRecipientLookup.Visibility = Visibility.Visible;
                    break;
                case "Reporting":
                    PanelReporting.Visibility = Visibility.Visible;
                    break;
                case "Snapshots":
                    PanelSnapshots.Visibility = Visibility.Visible;
                    RefreshSnapshotsList();
                    break;
                case "Settings":
                    PanelSettings.Visibility = Visibility.Visible;
                    break;
                default:
                    PanelMailboxes.Visibility = Visibility.Visible;
                    break;
            }
        }

        // --- Bulk list file import ---

        private async void ButtonBrowseSharedMailboxes_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxSharedMailboxes);
        private async void ButtonBrowseMailboxUsers_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxMailboxUsers);
        private async void ButtonBrowseResourceMailboxes_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxResourceMailboxes);
        private async void ButtonBrowseResourceUsers_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxResourceUsers);
        private async void ButtonBrowseGroups_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxGroups);
        private async void ButtonBrowseGroupUsers_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxGroupUsers);
        private async void ButtonBrowseBookingsUsers_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxBookingsUsers);
        private async void ButtonBrowseNewGroupMembers_Click(object sender, RoutedEventArgs e) => await ImportListFromFileAsync(TextBoxNewGroupMembers);

        private async Task ImportListFromFileAsync(TextBox targetBox)
            => await RunUiOperationAsync(() => ImportListCoreAsync(targetBox));

        private async Task ImportListCoreAsync(TextBox targetBox)
        {
            var picker = new FileOpenPicker();
            InitializeWithWindow.Initialize(picker, GetWindowHandle());
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            picker.FileTypeFilter.Add(".csv");
            picker.FileTypeFilter.Add(".txt");

            var file = await picker.PickSingleFileAsync();
            if (file == null)
            {
                return;
            }

            try
            {
                using var reader = new StringReader(await FileIO.ReadTextAsync(file));
                using var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader);
                parser.SetDelimiters(",", ";");
                parser.HasFieldsEnclosedInQuotes = true;
                var imported = new List<string>();
                while (!parser.EndOfData) imported.AddRange(parser.ReadFields() ?? Array.Empty<string>());
                var values = imported.Select(value => value.Trim()).Where(value => !string.IsNullOrEmpty(value)).ToArray();

                if (values.Length == 0)
                {
                    Logger.Log($"No values found in '{file.Name}'.", LogType.Warning);
                    return;
                }

                var combined = string.IsNullOrWhiteSpace(targetBox.Text)
                    ? string.Join(Environment.NewLine, values)
                    : targetBox.Text.TrimEnd() + Environment.NewLine + string.Join(Environment.NewLine, values);

                targetBox.Text = combined;
                Logger.Log($"Imported {values.Length} value(s) from '{file.Name}'.", LogType.Success);
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to read '{file.Name}': {ex.Message}", LogType.Error);
            }
        }

        private void UpdateConnectionStatus()
        {
            var exoStatus = _exo.IsConnected ? $"EXO: {_exo.ConnectedUserPrincipalName}" : "EXO: Not connected";
            var graphStatus = _authService.IsGraphConnected ? $"Graph: {_authService.ConnectedUser}" : "Graph: Not connected";
            TextBlockConnectionStatus.Text = $"{exoStatus} | {graphStatus}";
            ButtonConnectGraph.IsEnabled = _exo.IsConnected && !_authService.IsGraphConnected;
            ButtonConnectExo.IsEnabled = !_exo.IsConnected;
        }

        // --- Connection handlers ---

        private async void ButtonConnectExo_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(ConnectExoAsync);

        private async Task ConnectExoAsync()
        {
            ButtonConnectExo.IsEnabled = false;
            await _exo.ConnectAsync(CheckBoxExoBrowserSignIn.IsChecked == true);
            UpdateConnectionStatus();
            await RefreshAcceptedDomainsAsync();
        }

        private async void ButtonConnectGraph_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(ConnectGraphAsync);

        private async Task ConnectGraphAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to Exchange Online first.", "Connection Order");
                return;
            }

            ButtonConnectGraph.IsEnabled = false;
            if (!string.Equals(_exo.ConnectedTenantId, _config.Settings.GraphApi.TenantId, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(_exo.ConnectedTenantId))
            {
                await ShowMessageAsync("The EXO tenant must match the Graph TenantId in Settings. Disconnect and reconnect to the intended tenant.", "Tenant Mismatch");
                UpdateConnectionStatus();
                return;
            }
            var connected = await _authService.ConnectGraphAsync();
            if (connected)
            {
                try
                {
                    _graphService.InitializeGraphClient();
                    await _graphService.VerifyConnectionAsync();
                    await _exo.GetAcceptedDomainsAsync();
                    Logger.Log("Exchange Online and Microsoft Graph read checks passed with both sessions connected.", LogType.Success);
                }
                catch
                {
                    await _authService.DisconnectGraphAsync();
                    UpdateConnectionStatus();
                    Logger.Log("Combined EXO/Graph connection check failed. Graph was disconnected; verify Exchange access before continuing.", LogType.Error);
                    throw;
                }
            }
            UpdateConnectionStatus();
        }

        private async void ButtonDisconnect_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(DisconnectAsync);

        private async Task DisconnectAsync()
        {
            await _exo.DisconnectAsync();
            await _authService.DisconnectGraphAsync();
            _loadedGroupSettingsIdentity = null;
            UpdateConnectionStatus();
        }

        // --- Mailbox permission handlers ---

        private async void ButtonValidateMailboxes_Click(object sender, RoutedEventArgs e) =>
            await RunUiOperationAsync(() => ValidateTargetsAndUsersAsync(TextBoxSharedMailboxes, TextBoxMailboxUsers, PermissionTargetType.Mailbox));

        private async void ButtonValidateResources_Click(object sender, RoutedEventArgs e) =>
            await RunUiOperationAsync(() => ValidateTargetsAndUsersAsync(TextBoxResourceMailboxes, TextBoxResourceUsers, PermissionTargetType.Resource));

        private async Task ValidateTargetsAndUsersAsync(TextBox targetsBox, TextBox usersBox, PermissionTargetType targetType)
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Please connect to Exchange Online first.", "EXO Not Connected");
                return;
            }

            var targets = InputParsingHelpers.ConvertToInputList(targetsBox.Text);
            var users = InputParsingHelpers.ConvertToInputList(usersBox.Text);

            if (targets.Length == 0 && users.Length == 0)
            {
                await ShowMessageAsync("Enter at least one target and/or user to validate.", "Missing Input");
                return;
            }

            var targetLabel = targetType == PermissionTargetType.Resource ? "Resource" : "Mailbox";
            Logger.Log($"--- Validating {targetLabel} and User Input ---");

            var invalidTargets = new List<string>();
            foreach (var target in targets)
            {
                var exists = await _exo.MailboxExistsAsync(target, resourceOnly: targetType == PermissionTargetType.Resource);
                if (exists)
                {
                    Logger.Log($"{targetLabel} '{target}': OK", LogType.Success);
                }
                else
                {
                    Logger.Log($"{targetLabel} '{target}': NOT FOUND", LogType.Error);
                    invalidTargets.Add(target);
                }
            }

            var invalidUsers = new List<string>();
            foreach (var user in users)
            {
                var recipient = await _exo.GetRecipientAsync(user);
                if (recipient != null)
                {
                    Logger.Log($"User '{user}': OK", LogType.Success);
                }
                else
                {
                    Logger.Log($"User '{user}': NOT FOUND", LogType.Error);
                    invalidUsers.Add(user);
                }
            }

            if (invalidTargets.Count == 0 && invalidUsers.Count == 0)
            {
                Logger.Log("Validation complete. All entries are valid.", LogType.Success);
                await ShowMessageAsync("All targets and users are valid.", "Validation Passed");
            }
            else
            {
                var summary = new List<string>();
                if (invalidTargets.Count > 0) summary.Add($"{invalidTargets.Count} invalid {targetLabel.ToLowerInvariant()}(es): {string.Join(", ", invalidTargets)}");
                if (invalidUsers.Count > 0) summary.Add($"{invalidUsers.Count} invalid user(s): {string.Join(", ", invalidUsers)}");
                var message = string.Join(Environment.NewLine, summary);
                Logger.Log($"Validation complete with issues: {message}", LogType.Warning);
                await ShowMessageAsync(message, "Validation Issues Found");
            }
        }

        private async void ButtonAddMailboxPermissions_Click(object sender, RoutedEventArgs e) =>
            await RunMailboxPermissionOperationAsync(PermissionOperationType.Add, PermissionTargetType.Mailbox,
                TextBoxSharedMailboxes, TextBoxMailboxUsers, CheckBoxFullAccess, CheckBoxSendAs, CheckBoxSendOnBehalf,
                ButtonAddMailboxPermissions, ButtonRemoveMailboxPermissions);

        private async void ButtonRemoveMailboxPermissions_Click(object sender, RoutedEventArgs e) =>
            await RunMailboxPermissionOperationAsync(PermissionOperationType.Remove, PermissionTargetType.Mailbox,
                TextBoxSharedMailboxes, TextBoxMailboxUsers, CheckBoxFullAccess, CheckBoxSendAs, CheckBoxSendOnBehalf,
                ButtonAddMailboxPermissions, ButtonRemoveMailboxPermissions);

        private async void ButtonAddResourcePermissions_Click(object sender, RoutedEventArgs e) =>
            await RunMailboxPermissionOperationAsync(PermissionOperationType.Add, PermissionTargetType.Resource,
                TextBoxResourceMailboxes, TextBoxResourceUsers, CheckBoxResourceFullAccess, CheckBoxResourceSendAs, CheckBoxResourceSendOnBehalf,
                ButtonAddResourcePermissions, ButtonRemoveResourcePermissions);

        private async void ButtonRemoveResourcePermissions_Click(object sender, RoutedEventArgs e) =>
            await RunMailboxPermissionOperationAsync(PermissionOperationType.Remove, PermissionTargetType.Resource,
                TextBoxResourceMailboxes, TextBoxResourceUsers, CheckBoxResourceFullAccess, CheckBoxResourceSendAs, CheckBoxResourceSendOnBehalf,
                ButtonAddResourcePermissions, ButtonRemoveResourcePermissions);

        private async Task RunMailboxPermissionOperationAsync(
            PermissionOperationType operationType, PermissionTargetType targetType,
            TextBox targetsBox, TextBox usersBox, CheckBox fullAccessBox, CheckBox sendAsBox, CheckBox sendOnBehalfBox,
            Button addButton, Button removeButton)
            => await RunUiOperationAsync(() => RunMailboxPermissionCoreAsync(operationType, targetType, targetsBox, usersBox, fullAccessBox, sendAsBox, sendOnBehalfBox, addButton, removeButton));

        private async Task RunMailboxPermissionCoreAsync(
            PermissionOperationType operationType, PermissionTargetType targetType,
            TextBox targetsBox, TextBox usersBox, CheckBox fullAccessBox, CheckBox sendAsBox, CheckBox sendOnBehalfBox,
            Button addButton, Button removeButton)
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Please connect to Exchange Online first.", "EXO Not Connected");
                return;
            }

            var targets = InputParsingHelpers.ConvertToInputList(targetsBox.Text);
            var users = InputParsingHelpers.ConvertToInputList(usersBox.Text);
            var permissions = new PermissionSelections
            {
                FullAccess = fullAccessBox.IsChecked == true,
                SendAs = sendAsBox.IsChecked == true,
                SendOnBehalf = sendOnBehalfBox.IsChecked == true
            };

            if (targets.Length == 0 || users.Length == 0 || (!permissions.FullAccess && !permissions.SendAs && !permissions.SendOnBehalf))
            {
                await ShowMessageAsync("Enter at least one target, one user, and select at least one permission.", "Missing Input");
                return;
            }

            if (operationType == PermissionOperationType.Remove)
            {
                var confirmed = await ShowConfirmAsync($"Remove selected permissions for {users.Length} user(s) from {targets.Length} target(s)?", "Confirm Removal");
                if (!confirmed)
                {
                    Logger.Log("Mailbox permission removal cancelled by user.", LogType.Warning);
                    return;
                }
            }

            addButton.IsEnabled = false;
            removeButton.IsEnabled = false;
            try
            {
                var results = await _mailboxPermissionService.InvokePermissionOperationAsync(operationType, targetType, targets, users, permissions);
                if (results.Count == targets.Length * users.Length && results.All(result => result.Statuses.Count > 0 && result.Statuses.All(IsVerifiedStatus)))
                    await TryAutoCloseServiceNowTicketAsync();
                targetsBox.Text = string.Empty;
                usersBox.Text = string.Empty;
            }
            finally
            {
                addButton.IsEnabled = true;
                removeButton.IsEnabled = true;
            }
        }

        // --- Group membership handlers ---

        private async void ButtonAddMembership_Click(object sender, RoutedEventArgs e) =>
            await RunGroupMembershipOperationAsync(PermissionOperationType.Add);

        private async void ButtonRemoveMembership_Click(object sender, RoutedEventArgs e) =>
            await RunGroupMembershipOperationAsync(PermissionOperationType.Remove);

        private async Task RunGroupMembershipOperationAsync(PermissionOperationType operationType)
            => await RunUiOperationAsync(() => RunGroupMembershipCoreAsync(operationType));

        private async Task RunGroupMembershipCoreAsync(PermissionOperationType operationType)
        {
            var exoConnected = _exo.IsConnected;
            var graphConnected = _authService.IsGraphConnected;
            if (!exoConnected && !graphConnected)
            {
                await ShowMessageAsync("Please connect to Exchange Online and/or Microsoft Graph first.", "Connection Missing");
                return;
            }

            var groups = InputParsingHelpers.ConvertToInputList(TextBoxGroups.Text);
            var users = InputParsingHelpers.ConvertToInputList(TextBoxGroupUsers.Text);
            var roles = new GroupRoleSelections
            {
                Member = CheckBoxMember.IsChecked == true,
                Owner = CheckBoxOwner.IsChecked == true
            };

            if (groups.Length == 0 || users.Length == 0 || (!roles.Member && !roles.Owner))
            {
                await ShowMessageAsync("Enter at least one group, one user, and select at least one role.", "Missing Input");
                return;
            }

            if (operationType == PermissionOperationType.Remove)
            {
                var confirmed = await ShowConfirmAsync($"Remove selected role(s) for {users.Length} user(s) from {groups.Length} group(s)?", "Confirm Removal");
                if (!confirmed)
                {
                    Logger.Log("Group role removal cancelled by user.", LogType.Warning);
                    return;
                }
            }

            ButtonAddMembership.IsEnabled = false;
            ButtonRemoveMembership.IsEnabled = false;
            try
            {
                var results = await _groupMembershipService.InvokeGroupMembershipOperationAsync(operationType, groups, users, roles, exoConnected, graphConnected);
                if (results.Count == groups.Length * users.Length && results.All(result => result.Statuses.Count > 0 && result.Statuses.All(IsVerifiedStatus)))
                    await TryAutoCloseServiceNowTicketAsync();
                TextBoxGroups.Text = string.Empty;
                TextBoxGroupUsers.Text = string.Empty;
            }
            finally
            {
                ButtonAddMembership.IsEnabled = true;
                ButtonRemoveMembership.IsEnabled = true;
            }
        }

        // --- Group Settings handlers (Distribution Groups / Mail-Enabled Security Groups only) ---

        private async void ButtonLoadGroupSettings_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(LoadGroupSettingsAsync);

        private async Task LoadGroupSettingsAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            var identity = TextBoxGroupSettingsIdentity.Text.Trim();
            if (string.IsNullOrEmpty(identity))
            {
                await ShowMessageAsync("Enter a group identity.", "Missing Input");
                return;
            }

            ButtonLoadGroupSettings.IsEnabled = false;
            _loadedGroupSettingsIdentity = null;
            try
            {
                var snapshot = await _groupSettingsService.LoadSettingsAsync(identity);
                if (snapshot == null)
                {
                    return;
                }

                RadioDeliveryInternalAndExternal.IsChecked = snapshot.AllowExternalSenders;
                RadioDeliveryInternalOnly.IsChecked = !snapshot.AllowExternalSenders;
                TextBoxDeliverySpecifiedSenders.Text = string.Join(Environment.NewLine, snapshot.SpecifiedSenders);

                TextBoxDelegatesSendAs.Text = string.Join(Environment.NewLine, snapshot.SendAsDelegates);
                TextBoxDelegatesSendOnBehalf.Text = string.Join(Environment.NewLine, snapshot.SendOnBehalfDelegates);

                CheckBoxRequireModeratorApproval.IsChecked = snapshot.RequireModeratorApproval;
                TextBoxModerators.Text = string.Join(Environment.NewLine, snapshot.Moderators);
                TextBoxBypassModerationSenders.Text = string.Join(Environment.NewLine, snapshot.BypassModerationSenders);
                RadioNotifyOnlySender.IsChecked = string.Equals(snapshot.NotifySenderMode, "Always", StringComparison.OrdinalIgnoreCase);
                RadioNotifyOnlyInternal.IsChecked = string.Equals(snapshot.NotifySenderMode, "Internal", StringComparison.OrdinalIgnoreCase);
                RadioNotifyNever.IsChecked = string.Equals(snapshot.NotifySenderMode, "Never", StringComparison.OrdinalIgnoreCase);

                RadioGroupSettingsJoinOpen.IsChecked = string.Equals(snapshot.JoinRestriction, "Open", StringComparison.OrdinalIgnoreCase);
                RadioGroupSettingsJoinClosed.IsChecked = string.Equals(snapshot.JoinRestriction, "Closed", StringComparison.OrdinalIgnoreCase);
                RadioGroupSettingsJoinApproval.IsChecked = string.Equals(snapshot.JoinRestriction, "ApprovalRequired", StringComparison.OrdinalIgnoreCase);
                RadioGroupSettingsLeaveOpen.IsChecked = string.Equals(snapshot.DepartRestriction, "Open", StringComparison.OrdinalIgnoreCase);
                RadioGroupSettingsLeaveClosed.IsChecked = string.Equals(snapshot.DepartRestriction, "Closed", StringComparison.OrdinalIgnoreCase);
                if (string.Equals(identity, TextBoxGroupSettingsIdentity.Text.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    _loadedGroupSettingsIdentity = identity;
                }
            }
            finally
            {
                ButtonLoadGroupSettings.IsEnabled = true;
            }
        }

        private bool TryGetGroupSettingsIdentity(out string identity)
        {
            identity = TextBoxGroupSettingsIdentity.Text.Trim();
            return !string.IsNullOrEmpty(identity)
                && string.Equals(identity, _loadedGroupSettingsIdentity, StringComparison.OrdinalIgnoreCase);
        }

        private async void ButtonSaveDeliveryManagement_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(SaveDeliveryManagementAsync);

        private async Task SaveDeliveryManagementAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }
            if (!TryGetGroupSettingsIdentity(out var identity))
            {
                await ShowMessageAsync("Enter a group identity and load settings first.", "Missing Input");
                return;
            }

            var allowExternalSenders = RadioDeliveryInternalAndExternal.IsChecked == true;
            var specifiedSenders = InputParsingHelpers.ConvertToInputList(TextBoxDeliverySpecifiedSenders.Text);

            ButtonSaveDeliveryManagement.IsEnabled = false;
            try
            {
                await _groupSettingsService.SaveDeliveryManagementAsync(identity, allowExternalSenders, specifiedSenders);
            }
            finally
            {
                ButtonSaveDeliveryManagement.IsEnabled = true;
            }
        }

        private async void ButtonSaveDelegates_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(SaveDelegatesAsync);

        private async Task SaveDelegatesAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }
            if (!TryGetGroupSettingsIdentity(out var identity))
            {
                await ShowMessageAsync("Enter a group identity and load settings first.", "Missing Input");
                return;
            }

            var sendAsDelegates = InputParsingHelpers.ConvertToInputList(TextBoxDelegatesSendAs.Text);
            var sendOnBehalfDelegates = InputParsingHelpers.ConvertToInputList(TextBoxDelegatesSendOnBehalf.Text);

            if (!await ShowConfirmAsync($"Replace delegates on '{identity}' with the displayed lists? Delegates omitted from these lists will be removed.", "Confirm Delegate Changes")) return;

            ButtonSaveDelegates.IsEnabled = false;
            try
            {
                await _groupSettingsService.SaveDelegatesAsync(identity, sendAsDelegates, sendOnBehalfDelegates);
            }
            finally
            {
                ButtonSaveDelegates.IsEnabled = true;
            }
        }

        private async void ButtonSaveMessageApproval_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(SaveMessageApprovalAsync);

        private async Task SaveMessageApprovalAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }
            if (!TryGetGroupSettingsIdentity(out var identity))
            {
                await ShowMessageAsync("Enter a group identity and load settings first.", "Missing Input");
                return;
            }

            var requireModeratorApproval = CheckBoxRequireModeratorApproval.IsChecked == true;
            var moderators = InputParsingHelpers.ConvertToInputList(TextBoxModerators.Text);
            var bypassSenders = InputParsingHelpers.ConvertToInputList(TextBoxBypassModerationSenders.Text);
            var notifySenderMode = RadioNotifyOnlyInternal.IsChecked == true ? "Internal" : RadioNotifyNever.IsChecked == true ? "Never" : "Always";

            ButtonSaveMessageApproval.IsEnabled = false;
            try
            {
                await _groupSettingsService.SaveMessageApprovalAsync(identity, requireModeratorApproval, moderators, bypassSenders, notifySenderMode);
            }
            finally
            {
                ButtonSaveMessageApproval.IsEnabled = true;
            }
        }

        private async void ButtonSaveMembershipApproval_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(SaveMembershipApprovalAsync);

        private async Task SaveMembershipApprovalAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }
            if (!TryGetGroupSettingsIdentity(out var identity))
            {
                await ShowMessageAsync("Enter a group identity and load settings first.", "Missing Input");
                return;
            }

            var joinRestriction = RadioGroupSettingsJoinClosed.IsChecked == true ? "Closed" : RadioGroupSettingsJoinApproval.IsChecked == true ? "ApprovalRequired" : "Open";
            var departRestriction = RadioGroupSettingsLeaveClosed.IsChecked == true ? "Closed" : "Open";

            ButtonSaveMembershipApproval.IsEnabled = false;
            try
            {
                await _groupSettingsService.SaveMembershipApprovalAsync(identity, joinRestriction, departRestriction);
            }
            finally
            {
                ButtonSaveMembershipApproval.IsEnabled = true;
            }
        }

        // --- Bookings handler ---

        private async void ButtonEnableBookings_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(EnableBookingsAsync);

        private async Task EnableBookingsAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("EXO connection required.", "Connection Missing");
                return;
            }
            if (!_authService.IsGraphConnected)
            {
                await ShowMessageAsync("Graph connection required.", "Connection Missing");
                return;
            }

            var users = InputParsingHelpers.ConvertToInputList(TextBoxBookingsUsers.Text);
            if (users.Length == 0)
            {
                await ShowMessageAsync("Enter at least one user.", "Missing Input");
                return;
            }

            ButtonEnableBookings.IsEnabled = false;
            try
            {
                await _bookingsService.EnableBookingsAccessAsync(users);
                TextBoxBookingsUsers.Text = string.Empty;
            }
            finally
            {
                ButtonEnableBookings.IsEnabled = true;
            }
        }

        // --- Recipient lookup handler ---

        private async void ButtonCheckRecipient_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(CheckRecipientAsync);

        private async Task CheckRecipientAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            ButtonCheckRecipient.IsEnabled = false;
            TextBoxRecipientCheckResult.Text = "Checking...";
            try
            {
                var result = await _recipientLookupService.CheckRecipientTypeAsync(TextBoxRecipientCheck.Text.Trim());
                TextBoxRecipientCheckResult.Text = result.DisplayType;
            }
            finally
            {
                ButtonCheckRecipient.IsEnabled = true;
            }
        }

        // --- Reporting ---

        private ReportTargetInfo? _validatedReportTarget;

        private void TextBoxReportingIdentity_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Any edit to the identity invalidates the previous Validate result, so the user can't
            // run a report against a stale/mismatched type confirmation.
            _validatedReportTarget = null;
            ButtonGenerateObjectReport.IsEnabled = false;
            ButtonGetMailboxDelegates.IsEnabled = false;
            TextBlockReportingValidation.Text = "Enter an identity and click Validate to confirm its type before running a report.";
        }

        private async void ButtonValidateReportingIdentity_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(ValidateReportingIdentityAsync);

        private async Task ValidateReportingIdentityAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            var identity = TextBoxReportingIdentity.Text.Trim();
            if (string.IsNullOrEmpty(identity))
            {
                await ShowMessageAsync("Enter a group, distribution list, or mailbox identity.", "Input Missing");
                return;
            }

            ButtonValidateReportingIdentity.IsEnabled = false;
            ButtonGenerateObjectReport.IsEnabled = false;
            ButtonGetMailboxDelegates.IsEnabled = false;
            _validatedReportTarget = null;
            TextBlockReportingValidation.Text = "Validating...";
            try
            {
                var target = await _reportingService.ClassifyReportTargetAsync(identity);
                _validatedReportTarget = target;

                switch (target.Kind)
                {
                    case ReportTargetKind.Group:
                        ButtonGenerateObjectReport.IsEnabled = true;
                        TextBlockReportingValidation.Text = $"'{target.Name}' ({target.PrimarySmtpAddress}) is a {target.FriendlyType}. Use 'Generate Membership/Permissions Report'.";
                        break;
                    case ReportTargetKind.Mailbox:
                        ButtonGetMailboxDelegates.IsEnabled = true;
                        TextBlockReportingValidation.Text = $"'{target.Name}' ({target.PrimarySmtpAddress}) is a {target.FriendlyType}. Use 'Get Mailbox Delegates'.";
                        break;
                    default:
                        TextBlockReportingValidation.Text = $"'{target.Name}' ({target.PrimarySmtpAddress}) is a '{target.FriendlyType}', which is not supported for reporting.";
                        Logger.Log($"Reporting validation: '{identity}' resolved to unsupported type '{target.RecipientTypeDetails}'.", LogType.Warning);
                        break;
                }

                Logger.Log($"Reporting validation: '{identity}' is a {target.FriendlyType}.", LogType.Success);
            }
            catch (Exception ex)
            {
                TextBlockReportingValidation.Text = $"Validation failed: {ex.Message}";
                Logger.Log($"Reporting validation failed for '{identity}': {ex.Message}", LogType.Error);
                await ShowMessageAsync(ex.Message, "Validation Failed");
            }
            finally
            {
                ButtonValidateReportingIdentity.IsEnabled = true;
            }
        }

        private async void ButtonGenerateObjectReport_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(GenerateObjectReportAsync);

        private async Task GenerateObjectReportAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            var identity = TextBoxReportingIdentity.Text.Trim();
            if (string.IsNullOrEmpty(identity))
            {
                await ShowMessageAsync("Enter a group, distribution list, or mailbox identity.", "Input Missing");
                return;
            }

            if (_validatedReportTarget == null || _validatedReportTarget.Kind != ReportTargetKind.Group)
            {
                await ShowMessageAsync("Click Validate first to confirm this identity is a group or distribution list.", "Not Validated");
                return;
            }

            ButtonGenerateObjectReport.IsEnabled = false;
            ButtonGetMailboxDelegates.IsEnabled = false;
            ButtonExportReportCsv.IsEnabled = false;
            try
            {
                Logger.Log($"Resolving recipient '{identity}'...");
                var rows = await _reportingService.GenerateObjectReportAsync(identity);
                _reportResults.Clear();
                foreach (var row in rows)
                {
                    _reportResults.Add(row);
                }

                if (rows.Count == 0)
                {
                    Logger.Log($"No members or permissions found for '{identity}'.", LogType.Warning);
                }
                else
                {
                    Logger.Log($"Report complete: {rows.Count} row(s) for '{identity}'.", LogType.Success);
                    ButtonExportReportCsv.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Report generation failed for '{identity}': {ex.Message}", LogType.Error);
                await ShowMessageAsync(ex.Message, "Report Generation Failed");
            }
            finally
            {
                ButtonGenerateObjectReport.IsEnabled = _validatedReportTarget?.Kind == ReportTargetKind.Group;
                ButtonGetMailboxDelegates.IsEnabled = _validatedReportTarget?.Kind == ReportTargetKind.Mailbox;
            }
        }

        private async void ButtonGetMailboxDelegates_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(GetMailboxDelegatesAsync);

        private async Task GetMailboxDelegatesAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            var identity = TextBoxReportingIdentity.Text.Trim();
            if (string.IsNullOrEmpty(identity))
            {
                await ShowMessageAsync("Enter the mailbox or resource email address.", "Input Missing");
                return;
            }

            if (_validatedReportTarget == null || _validatedReportTarget.Kind != ReportTargetKind.Mailbox)
            {
                await ShowMessageAsync("Click Validate first to confirm this identity is a mailbox.", "Not Validated");
                return;
            }

            ButtonGenerateObjectReport.IsEnabled = false;
            ButtonGetMailboxDelegates.IsEnabled = false;
            ButtonExportReportCsv.IsEnabled = false;
            try
            {
                Logger.Log($"Getting delegates for '{identity}'...");
                var rows = await _reportingService.GetMailboxDelegatesAsync(identity);
                _reportResults.Clear();
                foreach (var row in rows)
                {
                    _reportResults.Add(row);
                }

                if (rows.Count == 0)
                {
                    Logger.Log($"No delegates found for '{identity}'.", LogType.Warning);
                }
                else
                {
                    Logger.Log($"Delegate report complete: {rows.Count} row(s) for '{identity}'.", LogType.Success);
                    ButtonExportReportCsv.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Delegate lookup failed for '{identity}': {ex.Message}", LogType.Error);
                await ShowMessageAsync(ex.Message, "Delegate Lookup Failed");
            }
            finally
            {
                ButtonGenerateObjectReport.IsEnabled = _validatedReportTarget?.Kind == ReportTargetKind.Group;
                ButtonGetMailboxDelegates.IsEnabled = _validatedReportTarget?.Kind == ReportTargetKind.Mailbox;
            }
        }

        private async void ButtonExportReportCsv_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(ExportReportCsvAsync);

        private async Task ExportReportCsvAsync()
        {
            if (_reportResults.Count == 0)
            {
                return;
            }

            var savePicker = new FileSavePicker();
            InitializeWithWindow.Initialize(savePicker, GetWindowHandle());
            savePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            savePicker.FileTypeChoices.Add("CSV File", new List<string> { ".csv" });
            savePicker.SuggestedFileName = $"EXOKit_Report_{DateTime.Now:yyyyMMdd_HHmmss}";

            var file = await savePicker.PickSaveFileAsync();
            if (file == null)
            {
                return;
            }

            try
            {
                var lines = new List<string> { "ObjectName,ObjectPrimarySmtpAddress,ObjectType,MemberOrDelegate,RoleOrPermission" };
                lines.AddRange(_reportResults.Select(r => string.Join(",",
                    CsvEscape(r.ObjectName), CsvEscape(r.ObjectPrimarySmtpAddress), CsvEscape(r.ObjectType), CsvEscape(r.MemberOrDelegate), CsvEscape(r.RoleOrPermission))));

                await FileIO.WriteLinesAsync(file, lines);
                Logger.Log($"Report exported to '{file.Path}'.", LogType.Success);
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to export report: {ex.Message}", LogType.Error);
            }
        }

        private static string CsvEscape(string value)
        {
            if (!string.IsNullOrEmpty(value) && ("=+-@".Contains(value.TrimStart().FirstOrDefault()) || value[0] is '\t' or '\r' or '\n')) value = "'" + value;
            if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        // --- Snapshots ---

        private void RefreshSnapshotsList()
        {
            _snapshotItems.Clear();
            foreach (var item in _snapshotService.ListSnapshotItems())
            {
                _snapshotItems.Add(item);
            }

            ButtonRestoreSnapshot.IsEnabled = false;
            ButtonDeleteSnapshot.IsEnabled = false;
            TextBlockSnapshotsStatus.Text = _snapshotItems.Count == 0
                ? "No snapshots found yet. Snapshots are created automatically before group membership/owner or mailbox permission removals."
                : $"{_snapshotItems.Count} snapshot(s) found.";
        }

        private void ButtonRefreshSnapshots_Click(object sender, RoutedEventArgs e)
        {
            RefreshSnapshotsList();
        }

        private void ListViewSnapshots_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ButtonRestoreSnapshot.IsEnabled = ListViewSnapshots.SelectedItem is SnapshotListItem;
            ButtonDeleteSnapshot.IsEnabled = ListViewSnapshots.SelectedItem is SnapshotListItem;
        }

        private async void ButtonDeleteSnapshot_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(DeleteSnapshotAsync);

        private async Task DeleteSnapshotAsync()
        {
            if (ListViewSnapshots.SelectedItem is not SnapshotListItem selected)
            {
                return;
            }

            var confirmed = await ShowConfirmAsync(
                $"Delete this snapshot?\n\n{selected.Description}\n\nThis cannot be undone and the removed item(s) will no longer be restorable from EXOKit.",
                "Delete Snapshot");
            if (!confirmed)
            {
                return;
            }

            try
            {
                if (_snapshotService.DeleteSnapshot(selected.Record))
                {
                    Logger.Log($"Snapshot deleted: '{selected.Description}'.", LogType.Success);
                }
                else
                {
                    Logger.Log($"Snapshot file could not be found on disk for '{selected.Description}'.", LogType.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to delete snapshot: {ex.Message}", LogType.Error);
                await ShowMessageAsync(ex.Message, "Delete Snapshot Failed");
            }
            finally
            {
                RefreshSnapshotsList();
            }
        }

        private async void ButtonClearAllSnapshots_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(ClearAllSnapshotsAsync);

        private async Task ClearAllSnapshotsAsync()
        {
            if (_snapshotItems.Count == 0)
            {
                return;
            }

            var confirmed = await ShowConfirmAsync(
                $"Delete all {_snapshotItems.Count} snapshot(s)?\n\nThis cannot be undone and none of the removed items will be restorable from EXOKit afterward.",
                "Clear All Snapshots");
            if (!confirmed)
            {
                return;
            }

            try
            {
                var count = _snapshotService.DeleteAllSnapshots();
                Logger.Log($"Cleared {count} snapshot(s).", LogType.Success);
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to clear snapshots: {ex.Message}", LogType.Error);
                await ShowMessageAsync(ex.Message, "Clear All Snapshots Failed");
            }
            finally
            {
                RefreshSnapshotsList();
            }
        }

        private void ButtonOpenSnapshotsFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Directory.CreateDirectory(_snapshotService.SnapshotsDirectory);

                // Passing the directory path directly as ProcessStartInfo.FileName with
                // UseShellExecute=true relies on the shell resolving a folder as its own "verb", which
                // can fail with "This location is not available" for MSIX-packaged apps (the per-user
                // LocalApplicationData path is a virtualized package folder that Explorer can be picky
                // about opening directly). Launching explorer.exe with the path as an argument is the
                // more reliable way to open a folder window from a packaged app.
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{_snapshotService.SnapshotsDirectory}\"",
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to open Snapshots folder ('{_snapshotService.SnapshotsDirectory}'): {ex.Message}", LogType.Error);
            }
        }

        private async void ButtonRestoreSnapshot_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(RestoreSnapshotAsync);

        private async Task RestoreSnapshotAsync()
        {
            if (!_exo.IsConnected) throw new InvalidOperationException("Connect Exchange Online before restoring a snapshot.");
            if (ListViewSnapshots.SelectedItem is not SnapshotListItem selected)
            {
                return;
            }

            var confirmed = await ShowConfirmAsync(
                $"Restore {selected.ItemCount} item(s) from this snapshot?\n\n{selected.Description}\n\nThis will re-add the removed member(s)/owner(s)/permission(s).",
                "Restore Snapshot");
            if (!confirmed)
            {
                return;
            }

            ButtonRestoreSnapshot.IsEnabled = false;
            TextBlockSnapshotsStatus.Text = "Restoring snapshot...";

            try
            {
                var results = await _snapshotRestoreService.RestoreAsync(selected.Record);
                var succeeded = results.Count(r => r.Success);
                var failed = results.Count - succeeded;

                TextBlockSnapshotsStatus.Text = failed == 0
                    ? $"Restore complete: {succeeded} item(s) restored successfully."
                    : $"Restore finished with issues: {succeeded} succeeded, {failed} failed. Check the log for details.";

                Logger.Log($"Snapshot restore complete for '{selected.Record.Target}': {succeeded} succeeded, {failed} failed.",
                    failed == 0 ? LogType.Success : LogType.Warning);
            }
            catch (Exception ex)
            {
                TextBlockSnapshotsStatus.Text = $"Restore failed: {ex.Message}";
                Logger.Log($"Snapshot restore failed: {ex.Message}", LogType.Error);
            }
            finally
            {
                ButtonRestoreSnapshot.IsEnabled = ListViewSnapshots.SelectedItem is SnapshotListItem;
            }
        }

        // --- Ticket notes ---

        private async void ButtonCopyTicketNotes_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(CopyTicketNotesAsync);

        private async Task CopyTicketNotesAsync()
        {
            var count = Logger.CopyLastTicketNotesToClipboard();
            if (count == 0)
            {
                await ShowMessageAsync("No ticket notes found in the current log.", "Copy Ticket Notes");
            }
            else
            {
                Logger.Log($"Copied {count} ticket note line(s) to clipboard.", LogType.Success);
            }
        }

        // --- Create Shared Mailbox ---

        private async void ButtonRefreshDomains_Click(object sender, RoutedEventArgs e) => await RunUiOperationAsync(RefreshAcceptedDomainsAsync);

        private async Task RefreshAcceptedDomainsAsync()
        {
            if (!_exo.IsConnected) return;

            try
            {
                var domains = await _exo.GetAcceptedDomainsAsync();
                var previouslySelected = ComboBoxNewMbDomain.SelectedItem as string;

                ComboBoxNewMbDomain.ItemsSource = domains;
                if (previouslySelected != null && domains.Contains(previouslySelected))
                {
                    ComboBoxNewMbDomain.SelectedItem = previouslySelected;
                }
                else if (domains.Length > 0)
                {
                    ComboBoxNewMbDomain.SelectedIndex = 0;
                }

                var previouslySelectedGroupDomain = ComboBoxNewGroupDomain.SelectedItem as string;
                ComboBoxNewGroupDomain.ItemsSource = domains;
                if (previouslySelectedGroupDomain != null && domains.Contains(previouslySelectedGroupDomain))
                {
                    ComboBoxNewGroupDomain.SelectedItem = previouslySelectedGroupDomain;
                }
                else if (domains.Length > 0)
                {
                    ComboBoxNewGroupDomain.SelectedIndex = 0;
                }

                Logger.Log($"Loaded {domains.Length} accepted domain(s).");
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to load accepted domains: {ex.Message}", LogType.Error);
            }
        }

        private async void ButtonCreateSharedMailbox_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(CreateSharedMailboxAsync);

        private async Task CreateSharedMailboxAsync()
        {
            if (!_exo.IsConnected)
            {
                await ShowMessageAsync("Connect to EXO first.", "EXO Not Connected");
                return;
            }

            var displayName = TextBoxNewMbDisplayName.Text.Trim();
            var alias = TextBoxNewMbAlias.Text.Trim();
            var domain = (ComboBoxNewMbDomain.SelectedItem as string)?.Trim();

            if (string.IsNullOrEmpty(displayName) || string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(domain))
            {
                await ShowMessageAsync("Display Name, Email Alias, and Domain are required.", "Input Missing");
                return;
            }

            var email = $"{alias}@{domain}";
            if (!InputParsingHelpers.IsEmailLikeValue(email))
            {
                await ShowMessageAsync($"'{email}' does not look like a valid email address.", "Invalid Email");
                return;
            }

            var request = new SharedMailboxCreationRequest
            {
                DisplayName = displayName,
                Email = email,
                Department = string.IsNullOrWhiteSpace(TextBoxNewMbDepartment.Text) ? null : TextBoxNewMbDepartment.Text.Trim(),
                EnableArchive = CheckBoxNewMbArchive.IsChecked == true,
                HideFromAddressLists = CheckBoxNewMbHideFromGal.IsChecked == true,
                RequireSenderAuthentication = CheckBoxNewMbRequireAuth.IsChecked == true,
                AcceptedSenders = InputParsingHelpers.ConvertToInputList(TextBoxNewMbAcceptedSenders.Text),
                BlockedSenders = InputParsingHelpers.ConvertToInputList(TextBoxNewMbBlockedSenders.Text),
                FullAccessUsers = InputParsingHelpers.ConvertToInputList(TextBoxNewMbFullAccess.Text),
                SendAsUsers = InputParsingHelpers.ConvertToInputList(TextBoxNewMbSendAs.Text),
                SendOnBehalfUsers = InputParsingHelpers.ConvertToInputList(TextBoxNewMbSendOnBehalf.Text)
            };

            var confirmed = await ShowConfirmAsync($"Create shared mailbox '{displayName}' ({email})?", "Confirm Mailbox Creation");
            if (!confirmed) return;

            ButtonCreateSharedMailbox.IsEnabled = false;
            try
            {
                var result = await _sharedMailboxService.CreateSharedMailboxAsync(request);
                if (result.Success)
                {
                    Logger.Log($"Shared mailbox '{email}' created successfully.", LogType.Success);
                    await ShowMessageAsync($"Shared mailbox '{email}' created successfully.", "Mailbox Created");
                }
                else
                {
                    await ShowMessageAsync($"Mailbox workflow did not fully complete for '{email}'. The mailbox may already exist; review the results before retrying.", "Incomplete Creation");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Shared mailbox creation failed: {ex.Message}", LogType.Error);
                await ShowMessageAsync($"Shared mailbox creation failed: {ex.Message}", "Error");
            }
            finally
            {
                ButtonCreateSharedMailbox.IsEnabled = true;
            }
        }

        // --- Create Group ---

        private void RadioNewGroupType_Checked(object sender, RoutedEventArgs e)
        {
            if (TextBlockNewGroupPrivacyLabel == null || PanelNewGroupPrivacy == null ||
                PanelNewGroupTeams == null || PanelNewGroupCommunication == null ||
                PanelNewGroupJoinLeave == null || PanelNewGroupApproval == null)
            {
                return;
            }

            var isM365 = RadioNewGroupM365?.IsChecked == true;
            var isSecurity = RadioNewGroupSecurity?.IsChecked == true;
            var isDistribution = RadioNewGroupDistribution?.IsChecked == true;

            TextBlockNewGroupPrivacyLabel.Visibility = isM365 ? Visibility.Visible : Visibility.Collapsed;
            PanelNewGroupPrivacy.Visibility = isM365 ? Visibility.Visible : Visibility.Collapsed;
            PanelNewGroupTeams.Visibility = isM365 ? Visibility.Visible : Visibility.Collapsed;

            PanelNewGroupCommunication.Visibility = (isDistribution || isSecurity) ? Visibility.Visible : Visibility.Collapsed;
            PanelNewGroupJoinLeave.Visibility = isDistribution ? Visibility.Visible : Visibility.Collapsed;
            PanelNewGroupApproval.Visibility = Visibility.Collapsed;
        }

        private async void ButtonRefreshGroupDomains_Click(object sender, RoutedEventArgs e) => await RunUiOperationAsync(RefreshAcceptedDomainsAsync);

        private async void ButtonCreateGroup_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(CreateGroupAsync);

        private async Task CreateGroupAsync()
        {
            var isM365 = RadioNewGroupM365.IsChecked == true;
            var isSecurity = RadioNewGroupSecurity.IsChecked == true;
            if (isM365 && CheckBoxNewGroupCreateTeam.IsChecked == true && !_authService.IsGraphConnected)
                throw new InvalidOperationException("Connect Microsoft Graph before creating a group with a Team.");

            if (!_exo.IsConnected)
            {
                var connectMessage = isM365
                    ? "Connect to EXO first. Microsoft 365 Groups are created via Exchange Online (New-UnifiedGroup)."
                    : "Connect to EXO first.";
                await ShowMessageAsync(connectMessage, "EXO Not Connected");
                return;
            }

            var displayName = TextBoxNewGroupDisplayName.Text.Trim();
            var alias = TextBoxNewGroupAlias.Text.Trim();
            var domain = (ComboBoxNewGroupDomain.SelectedItem as string)?.Trim();

            if (string.IsNullOrEmpty(displayName) || string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(domain))
            {
                await ShowMessageAsync("Display Name, Email Alias, and Domain are required.", "Input Missing");
                return;
            }

            var email = $"{alias}@{domain}";
            if (!InputParsingHelpers.IsEmailLikeValue(email))
            {
                await ShowMessageAsync($"'{email}' does not look like a valid email address.", "Invalid Email");
                return;
            }

            var owner = TextBoxNewGroupOwner.Text.Trim();
            if (string.IsNullOrEmpty(owner))
            {
                await ShowMessageAsync("At least one owner is required. Groups cannot be created without an owner.", "Owner Required");
                return;
            }

            var members = InputParsingHelpers.ConvertToInputList(TextBoxNewGroupMembers.Text);

            var request = new GroupCreationRequest
            {
                GroupKind = isM365 ? GroupKind.M365 : GroupKind.DistributionGroup,
                DisplayName = displayName,
                Alias = alias,
                Email = email,
                IsSecurityGroup = isSecurity,
                IsPrivate = RadioNewGroupPrivate.IsChecked == true,
                Owner = owner,
                Members = members,
                CreateTeam = CheckBoxNewGroupCreateTeam.IsChecked == true,
                AllowExternalSenders = CheckBoxNewGroupAllowExternalSenders.IsChecked == true,
                JoinRestriction = RadioNewGroupJoinClosed.IsChecked == true ? GroupJoinRestriction.Closed
                    : RadioNewGroupJoinApproval.IsChecked == true ? GroupJoinRestriction.ApprovalRequired
                    : GroupJoinRestriction.Open,
                DepartRestriction = RadioNewGroupLeaveClosed.IsChecked == true ? GroupDepartRestriction.Closed : GroupDepartRestriction.Open,
                RequireOwnerApprovalToJoin = CheckBoxNewGroupRequireApproval.IsChecked == true
            };

            var kindLabel = isM365 ? "Microsoft 365 Group" : (isSecurity ? "Mail-Enabled Security Group" : "Distribution Group");
            var confirmed = await ShowConfirmAsync($"Create {kindLabel} '{displayName}' ({email})?", "Confirm Group Creation");
            if (!confirmed) return;

            ButtonCreateGroup.IsEnabled = false;
            try
            {
                var result = await _groupCreationService.CreateGroupAsync(request);
                if (result.Success)
                {
                    Logger.Log($"{kindLabel} '{email}' created successfully.", LogType.Success);
                    await ShowMessageAsync($"{kindLabel} '{email}' created successfully.", "Group Created");
                }
                else
                {
                    await ShowMessageAsync($"Group workflow did not fully complete for '{email}'. The group may already exist; review the results before retrying.", "Incomplete Creation");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Group creation failed: {ex.Message}", LogType.Error);
                await ShowMessageAsync($"Group creation failed: {ex.Message}", "Error");
            }
            finally
            {
                ButtonCreateGroup.IsEnabled = true;
            }
        }

        // --- Settings ---

        private async void ButtonResumeTeam_Click(object sender, RoutedEventArgs args) => await RunUiOperationAsync(async () =>
        {
            if (!_exo.IsConnected || !_authService.IsGraphConnected || _exo.ConnectedTenantId != _authService.ConnectedTenantId)
                throw new InvalidOperationException("Connect EXO and Graph to the same tenant first.");
            var identity = TextBoxExistingTeamGroup.Text.Trim();
            if (string.IsNullOrEmpty(identity)) throw new InvalidOperationException("Enter the existing Microsoft 365 group email or object ID.");
            var groupId = await _graphService.ResolveM365GroupIdAsync(identity) ?? throw new InvalidOperationException("Microsoft 365 group not found.");
            if (!await ShowConfirmAsync($"Provision a Team for existing group '{identity}' ({groupId})?", "Resume Team Provisioning")) return;
            await _graphService.CreateTeamFromGroupAsync(groupId);
            Logger.Log($"Team is provisioned for group '{identity}' ({groupId}).", LogType.Success);
        });

        private void LoadSettingsIntoUi()
        {
            CheckBoxExoBrowserSignIn.IsChecked = _config.Settings.ExoUseBrowserSignIn;
            TextBoxSettingsBookingsGroup.Text = _config.Settings.LicenseGroups.Bookings.GroupName;
            TextBoxSettingsBookingsGroupId.Text = _config.Settings.LicenseGroups.Bookings.GroupId ?? string.Empty;
            TextBoxSettingsOwaPolicy.Text = _config.Settings.OwaPolicies.BookingsCreators;
            TextBoxSettingsGraphScopes.Text = string.Join(Environment.NewLine, _config.Settings.GraphApi.Scopes);
            TextBoxSettingsClientId.Text = _config.Settings.GraphApi.ClientId;
            TextBoxSettingsTenantId.Text = _config.Settings.GraphApi.TenantId;
            TextBoxSettingsUpdateFeed.Text = _config.Settings.UpdateFeedUrl;

            var serviceNow = _config.Settings.ServiceNow;
            CheckBoxSettingsServiceNowEnabled.IsChecked = serviceNow?.Enabled ?? false;
            TextBoxSettingsServiceNowInstanceUrl.Text = serviceNow?.InstanceUrl ?? string.Empty;
            TextBoxSettingsServiceNowKeyVaultUrl.Text = serviceNow?.KeyVaultUrl ?? string.Empty;
            TextBoxSettingsServiceNowSubscriptionId.Text = serviceNow?.SubscriptionId ?? string.Empty;
            TextBoxSettingsServiceNowUsernameSecret.Text = serviceNow?.UsernameSecretName ?? string.Empty;
            TextBoxSettingsServiceNowPasswordSecret.Text = serviceNow?.PasswordSecretName ?? string.Empty;
        }

        private async void ButtonSaveSettings_Click(object sender, RoutedEventArgs e)
            => await RunUiOperationAsync(SaveSettingsAsync);

        private async Task SaveSettingsAsync()
        {
            var bookingsGroupName = TextBoxSettingsBookingsGroup.Text.Trim();
            var bookingsGroupId = TextBoxSettingsBookingsGroupId.Text.Trim();
            if (!string.IsNullOrEmpty(bookingsGroupId) && !Guid.TryParse(bookingsGroupId, out _)) throw new InvalidOperationException("Bookings Group Object ID must be a GUID.");
            var owaPolicy = TextBoxSettingsOwaPolicy.Text.Trim();
            var scopes = InputParsingHelpers.ConvertToInputList(TextBoxSettingsGraphScopes.Text).ToList();
            var clientId = TextBoxSettingsClientId.Text.Trim();
            var tenantId = TextBoxSettingsTenantId.Text.Trim();
            if (!Guid.TryParse(clientId, out _) || !Guid.TryParse(tenantId, out _))
            {
                await ShowMessageAsync("Enter valid application and tenant GUIDs for the EXOKit app registration.", "Invalid Graph Configuration");
                return;
            }

            if ((string.IsNullOrEmpty(bookingsGroupName) && string.IsNullOrEmpty(bookingsGroupId)) || string.IsNullOrEmpty(owaPolicy) || scopes.Count == 0)
            {
                TextBlockSettingsStatus.Text = "Bookings Group Name, OWA Policy, and at least one Graph scope are required.";
                await ShowMessageAsync("Bookings Group Name, OWA Policy, and at least one Graph scope are required.", "Input Missing");
                return;
            }

            var serviceNowEnabled = CheckBoxSettingsServiceNowEnabled.IsChecked == true;
            var instanceUrl = TextBoxSettingsServiceNowInstanceUrl.Text.Trim();
            var keyVaultUrl = TextBoxSettingsServiceNowKeyVaultUrl.Text.Trim();
            var subscriptionId = TextBoxSettingsServiceNowSubscriptionId.Text.Trim();
            var usernameSecret = TextBoxSettingsServiceNowUsernameSecret.Text.Trim();
            var passwordSecret = TextBoxSettingsServiceNowPasswordSecret.Text.Trim();

            // Validate ServiceNow connectivity/credentials before persisting anything, so the user
            // finds out immediately if the instance URL, Key Vault, or secrets are wrong.
            if (serviceNowEnabled)
            {
                var candidateConfig = new ServiceNowConfig
                {
                    Enabled = true,
                    InstanceUrl = instanceUrl,
                    KeyVaultUrl = keyVaultUrl,
                    SubscriptionId = subscriptionId,
                    UsernameSecretName = usernameSecret,
                    PasswordSecretName = passwordSecret,
                    Table = _config.Settings.ServiceNow?.Table ?? "task",
                    TicketNumberField = _config.Settings.ServiceNow?.TicketNumberField ?? "number",
                    CloseStateField = _config.Settings.ServiceNow?.CloseStateField ?? "state",
                    CloseStateValue = _config.Settings.ServiceNow?.CloseStateValue ?? "3",
                    WorkNotesField = _config.Settings.ServiceNow?.WorkNotesField ?? "work_notes",
                    AdditionalCommentsField = _config.Settings.ServiceNow?.AdditionalCommentsField ?? "comments",
                    UpdateTable = _config.Settings.ServiceNow?.UpdateTable ?? string.Empty,
                    AdditionalFields = _config.Settings.ServiceNow?.AdditionalFields
                };

                TextBlockSettingsStatus.Text = "Validating ServiceNow connection...";
                ButtonSaveSettings.IsEnabled = false;
                try
                {
                    var testResult = await new ServiceNowService(candidateConfig, new GraphApiConfig { ClientId = clientId, TenantId = tenantId }).TestConnectionAsync();
                    if (!testResult.Success)
                    {
                        TextBlockSettingsStatus.Text = $"ServiceNow validation failed: {testResult.Message}";
                        await ShowMessageAsync(testResult.Message, "ServiceNow Validation Failed");
                        return;
                    }
                }
                finally
                {
                    ButtonSaveSettings.IsEnabled = true;
                }
            }

            var candidate = System.Text.Json.JsonSerializer.Deserialize<ToolConfig>(System.Text.Json.JsonSerializer.Serialize(_config))!;
            candidate.Settings.LicenseGroups.Bookings.GroupName = bookingsGroupName;
            candidate.Settings.ExoUseBrowserSignIn = CheckBoxExoBrowserSignIn.IsChecked == true;
            candidate.Settings.LicenseGroups.Bookings.GroupId = bookingsGroupId;
            candidate.Settings.OwaPolicies.BookingsCreators = owaPolicy;
            candidate.Settings.GraphApi.Scopes = scopes;
            candidate.Settings.GraphApi.ClientId = clientId;
            candidate.Settings.GraphApi.TenantId = tenantId;
            candidate.Settings.UpdateFeedUrl = TextBoxSettingsUpdateFeed.Text.Trim();
            if (!string.IsNullOrEmpty(candidate.Settings.UpdateFeedUrl)
                && (!Uri.TryCreate(candidate.Settings.UpdateFeedUrl, UriKind.Absolute, out var feedUri) || feedUri.Scheme != "https" || !string.IsNullOrEmpty(feedUri.UserInfo)))
                throw new InvalidOperationException("Update feed must be an HTTPS URL without embedded credentials.");

            if (serviceNowEnabled || !string.IsNullOrEmpty(instanceUrl) || !string.IsNullOrEmpty(keyVaultUrl))
            {
                var serviceNow = candidate.Settings.ServiceNow ??= new ServiceNowConfig();
                serviceNow.Enabled = serviceNowEnabled;
                serviceNow.InstanceUrl = instanceUrl;
                serviceNow.KeyVaultUrl = keyVaultUrl;
                serviceNow.SubscriptionId = subscriptionId;
                serviceNow.UsernameSecretName = usernameSecret;
                serviceNow.PasswordSecretName = passwordSecret;
            }
            else
            {
                candidate.Settings.ServiceNow = null;
            }

            ConfigService.Save(candidate);
            _config.Settings = candidate.Settings;
            await _authService.DisconnectGraphAsync();
            _authService = new AuthService(_config.Settings.GraphApi.Scopes.ToArray(), GetWindowHandle, clientId, tenantId);
            _graphService = new GraphService(_authService, _config.Settings.GraphApi.Scopes.ToArray());
            _groupMembershipService = new GroupMembershipService(_exo, _graphService, _snapshotService);
            _bookingsService = new BookingsService(_exo, _graphService, _config);
            _groupCreationService = new GroupCreationService(_exo, _graphService);
            _snapshotRestoreService = new SnapshotRestoreService(_exo, _graphService);
            UpdateConnectionStatus();
            _serviceNowService = _config.Settings.ServiceNow != null ? new ServiceNowService(_config.Settings.ServiceNow, _config.Settings.GraphApi) : null;

            try
            {
                TextBlockSettingsStatus.Text = $"Settings saved at {DateTime.Now:HH:mm:ss}.";
                Logger.Log("Settings saved to config.json.", LogType.Success);
            }
            catch (Exception ex)
            {
                TextBlockSettingsStatus.Text = $"Failed to save settings: {ex.Message}";
                Logger.Log($"Failed to save settings: {ex.Message}", LogType.Error);
                await ShowMessageAsync($"Failed to save settings: {ex.Message}", "Save Failed");
            }
        }

        // --- Dialog helpers ---

        private static bool IsVerifiedStatus(string status) =>
            status.EndsWith("(Added)", StringComparison.Ordinal)
            || status.EndsWith("(Removed)", StringComparison.Ordinal)
            || status.EndsWith("(Already Exists)", StringComparison.Ordinal);

        private async Task ShowMessageAsync(string message, string title)
        {
            var dialog = new ContentDialog
            {
                Title = title,
                Content = message,
                CloseButtonText = "OK",
                XamlRoot = Content.XamlRoot
            };
            await dialog.ShowAsync();
        }

        private async Task<bool> ShowConfirmAsync(string message, string title)
        {
            var dialog = new ContentDialog
            {
                Title = title,
                Content = message,
                PrimaryButtonText = "Yes",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Close,
                XamlRoot = Content.XamlRoot
            };
            var result = await dialog.ShowAsync();
            return result == ContentDialogResult.Primary;
        }
    }
}

