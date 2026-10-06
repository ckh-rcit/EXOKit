using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using EXOKit.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace EXOKit
{
    public sealed partial class MainWindow
    {
        private CalendarPermissionInfo? _calendarPermission;
        private MeetingOrganizerRequest? _validatedMeetingOrganizer;
        private string? _calendarTenantId;
        private readonly ObservableCollection<MeetingOrganizerMatch> _meetingMatches = new();
        private readonly ObservableCollection<CalendarTransferRecord> _calendarTransfers = new();

        private void InitializeCalendarControls()
        {
            ComboBoxCalendarRole.ItemsSource = ExoPowerShellService.CalendarPermissionRoles;
            ComboBoxCalendarRole.SelectedIndex = 0;
            TextBoxCalendarOwner.TextChanged += (_, _) => ResetCalendarPermission();
            TextBoxCalendarUser.TextChanged += (_, _) => ResetCalendarPermission();
            TextBoxCurrentOrganizer.TextChanged += (_, _) => ResetMeetingOrganizer();
            TextBoxNewOrganizer.TextChanged += (_, _) => ResetMeetingOrganizer();
            TextBoxMeetingSelector.TextChanged += (_, _) => ResetMeetingOrganizer();
            ComboBoxMeetingSelector.SelectionChanged += (_, _) =>
            {
                TextBoxMeetingSelector.Header = ComboBoxMeetingSelector.SelectedIndex == 0 ? "Series Event ID" : "Meeting subject";
                TextBoxMeetingSelector.Text = string.Empty;
                ResetMeetingOrganizer();
            };
            CheckBoxTransferDate.Checked += (_, _) => ResetMeetingOrganizer();
            CheckBoxTransferDate.Unchecked += (_, _) => ResetMeetingOrganizer();
            DatePickerTransferStart.MinDate = DateTimeOffset.Now.Date.AddDays(1);
            DatePickerTransferStart.DateChanged += (_, _) => ResetMeetingOrganizer();
            ListViewMeetingMatches.ItemsSource = _meetingMatches;
            ListViewCalendarTransfers.ItemsSource = _calendarTransfers;
        }

        private static void SetCalendarStatus(InfoBar status, string title, string message, InfoBarSeverity severity)
        {
            status.Title = title;
            status.Message = message;
            status.Severity = severity;
        }

        private void ResetCalendarPermission()
        {
            _calendarPermission = null;
            TextBlockCalendarPermission.Text = "Current: Not loaded";
            SetCalendarStatus(CalendarPermissionStatus, _exo.IsConnected ? "Not loaded" : "EXO not connected", "", InfoBarSeverity.Informational);
            UpdateCalendarControls();
        }

        private void ResetMeetingOrganizer()
        {
            _validatedMeetingOrganizer = null;
            _meetingMatches.Clear();
            MeetingMatchesPanel.Visibility = Visibility.Collapsed;
            SetCalendarStatus(MeetingOrganizerStatus, _exo.IsConnected ? "Not previewed" : "EXO not connected", "", InfoBarSeverity.Informational);
            UpdateCalendarControls();
        }

        private void UpdateCalendarConnectionState()
        {
            if (!_exo.IsConnected || (_calendarTenantId != null && !string.Equals(_calendarTenantId, _exo.ConnectedTenantId, StringComparison.OrdinalIgnoreCase)))
            {
                ResetCalendarPermission();
                ResetMeetingOrganizer();
                _calendarTransfers.Clear();
                _calendarTenantId = null;
            }
            else if (_calendarTenantId == null)
            {
                SetCalendarStatus(CalendarPermissionStatus, "Not loaded", "", InfoBarSeverity.Informational);
                SetCalendarStatus(MeetingOrganizerStatus, "Not previewed", "", InfoBarSeverity.Informational);
            }
            UpdateCalendarControls();
        }

        private void UpdateCalendarControls()
        {
            var ready = !_operationInProgress && _exo.IsConnected;
            ButtonLoadCalendarPermission.IsEnabled = ready && !string.IsNullOrWhiteSpace(TextBoxCalendarOwner.Text) && !string.IsNullOrWhiteSpace(TextBoxCalendarUser.Text);
            ButtonSetCalendarPermission.IsEnabled = ready && _calendarPermission != null;
            ButtonRemoveCalendarPermission.IsEnabled = ready && _calendarPermission?.HasEntry == true;
            DatePickerTransferStart.IsEnabled = !_operationInProgress && CheckBoxTransferDate.IsChecked == true;
            ButtonPreviewMeetingOrganizer.IsEnabled = ready && !string.IsNullOrWhiteSpace(TextBoxCurrentOrganizer.Text)
                && !string.IsNullOrWhiteSpace(TextBoxNewOrganizer.Text) && !string.IsNullOrWhiteSpace(TextBoxMeetingSelector.Text)
                && (CheckBoxTransferDate.IsChecked != true || DatePickerTransferStart.Date.HasValue);
            ButtonChangeMeetingOrganizer.IsEnabled = ready && _validatedMeetingOrganizer != null;
            ButtonExportCalendarTransfers.IsEnabled = !_operationInProgress && _calendarTransfers.Count > 0;
            ButtonClearCalendarTransfers.IsEnabled = !_operationInProgress && _calendarTransfers.Count > 0;
            TextBlockCalendarTransferCount.Text = $"Operation history: {_calendarTransfers.Count}";
        }

        private Task RunCalendarActionAsync(InfoBar status, Func<Task> action, bool readOnly) => RunUiOperationAsync(async () =>
        {
            try
            {
                if (!_exo.IsConnected) throw new InvalidOperationException("Connect to Exchange Online first.");
                _calendarTenantId = _exo.ConnectedTenantId;
                await action();
            }
            catch (Exception exception)
            {
                if (status == CalendarPermissionStatus) ResetCalendarPermission();
                else ResetMeetingOrganizer();
                SetCalendarStatus(status, readOnly ? "Preview or read incomplete" : "Result unconfirmed", exception.Message,
                    exception is OperationCanceledException ? InfoBarSeverity.Warning : InfoBarSeverity.Error);
                throw;
            }
            finally { UpdateCalendarControls(); }
        }, readOnly);

        private void ShowCalendarPermission(CalendarPermissionInfo permission)
        {
            _calendarPermission = permission;
            TextBlockCalendarPermission.Text = $"Calendar: {permission.FolderIdentity}\nUser: {permission.UserIdentity}\n"
                + (permission.HasEntry ? $"Explicit permission: {string.Join(", ", permission.AccessRights)}\nDelegate flags: {string.Join(", ", permission.SharingPermissionFlags)}"
                    : "No explicit permission entry. Default or group-based access may still apply.");
            if (permission.AccessRights.Length == 1 && ExoPowerShellService.CalendarPermissionRoles.Contains(permission.AccessRights[0]))
                ComboBoxCalendarRole.SelectedItem = permission.AccessRights[0];
        }

        private async void ButtonLoadCalendarPermission_Click(object sender, RoutedEventArgs args) =>
            await RunCalendarActionAsync(CalendarPermissionStatus, async () =>
            {
                ResetCalendarPermission();
                SetCalendarStatus(CalendarPermissionStatus, "Loading", "", InfoBarSeverity.Informational);
                var permission = await _exo.GetCalendarPermissionAsync(TextBoxCalendarOwner.Text.Trim(), TextBoxCalendarUser.Text.Trim());
                _exo.OperationCancellationToken.ThrowIfCancellationRequested();
                ShowCalendarPermission(permission);
                SetCalendarStatus(CalendarPermissionStatus, "Current entry loaded", "", InfoBarSeverity.Success);
                Logger.Log($"Calendar permission loaded: {permission.UserIdentity} on {permission.FolderIdentity}.");
            }, true);

        private async void ButtonSetCalendarPermission_Click(object sender, RoutedEventArgs args) => await ChangeCalendarPermissionAsync(false);

        private async void ButtonRemoveCalendarPermission_Click(object sender, RoutedEventArgs args) => await ChangeCalendarPermissionAsync(true);

        private Task ChangeCalendarPermissionAsync(bool remove) => RunCalendarActionAsync(CalendarPermissionStatus, async () =>
        {
            var loaded = _calendarPermission ?? throw new InvalidOperationException("Load the current permission first.");
            var role = ComboBoxCalendarRole.SelectedItem as string ?? throw new InvalidOperationException("Select a permission role.");
            var confirmation = remove
                ? $"Remove the explicit permission entry for {loaded.UserIdentity} on {loaded.FolderIdentity}?\n\nDefault or group-based access may still apply."
                : $"Set the explicit permission for {loaded.UserIdentity} on {loaded.FolderIdentity} to {role}?\n\nThis replaces the user's current folder permissions.";
            if (!await ShowConfirmAsync(confirmation, remove ? "Remove Calendar Entry" : "Set Calendar Permission")) return;
            _calendarPermission = null;
            SetCalendarStatus(CalendarPermissionStatus, "Applying and verifying", "", InfoBarSeverity.Informational);
            var result = remove
                ? await _exo.RemoveCalendarPermissionAsync(TextBoxCalendarOwner.Text.Trim(), TextBoxCalendarUser.Text.Trim(), loaded)
                : await _exo.SetCalendarPermissionAsync(TextBoxCalendarOwner.Text.Trim(), TextBoxCalendarUser.Text.Trim(), role, loaded);
            _exo.OperationCancellationToken.ThrowIfCancellationRequested();
            ShowCalendarPermission(result);
            SetCalendarStatus(CalendarPermissionStatus, remove ? "Entry removal verified" : "Permission verified", "", InfoBarSeverity.Success);
            Logger.WriteSummary(remove ? "Remove Calendar Entry" : "Set Calendar Permission", new Dictionary<string, object>
            {
                [result.UserIdentity] = new[] { $"Calendar: {result.FolderIdentity}", remove ? "Explicit entry removed and verified" : $"Permission: {role}; verified" }
            });
        }, false);

        private MeetingOrganizerRequest ReadMeetingOrganizerRequest()
        {
            if (CheckBoxTransferDate.IsChecked == true && !DatePickerTransferStart.Date.HasValue)
                throw new InvalidOperationException("Select a future start date.");
            return new MeetingOrganizerRequest(TextBoxCurrentOrganizer.Text.Trim(), TextBoxNewOrganizer.Text.Trim(),
                ComboBoxMeetingSelector.SelectedIndex == 0, TextBoxMeetingSelector.Text.Trim(),
                CheckBoxTransferDate.IsChecked == true ? DatePickerTransferStart.Date?.Date : null);
        }

        private void RecordCalendarTransfer(MeetingOrganizerRequest request, string status, string details)
        {
            _calendarTransfers.Add(new CalendarTransferRecord(DateTimeOffset.UtcNow, request.CurrentOrganizer, request.NewOrganizer,
                request.UseEventId ? "Event ID" : "Subject", request.Selector,
                request.TransferSeriesStartDate?.ToString("yyyy-MM-dd") ?? "Next instance", status, details));
        }

        private async void ButtonPreviewMeetingOrganizer_Click(object sender, RoutedEventArgs args) => await RunMeetingOrganizerAsync(true);

        private async void ButtonChangeMeetingOrganizer_Click(object sender, RoutedEventArgs args) => await RunMeetingOrganizerAsync(false);

        private Task RunMeetingOrganizerAsync(bool preview) => RunCalendarActionAsync(MeetingOrganizerStatus, async () =>
        {
            var request = preview ? ReadMeetingOrganizerRequest() : _validatedMeetingOrganizer
                ?? throw new InvalidOperationException("Preview the organizer change first.");
            if (!preview)
            {
                if (!await ShowConfirmAsync($"Current: {request.CurrentOrganizer}\nNew: {request.NewOrganizer}\n{(request.UseEventId ? "Series Event ID" : "Subject")}: {request.Selector}\nEffective: {request.TransferSeriesStartDate?.ToShortDateString() ?? "Next instance"}\n\n"
                    + "Only the default calendar is supported. The previous organizer is not retained as an attendee. External and on-premises attendees receive cancellation/invitation messages and must RSVP again.\n\n"
                    + "Teams ownership and file permissions do not transfer. The new organizer must replace Teams meeting details and check linked files. Earlier instances stay with the original organizer and cannot be transferred again.", "Change Meeting Organizer")) return;
            }
            _validatedMeetingOrganizer = null;
            _meetingMatches.Clear();
            MeetingMatchesPanel.Visibility = Visibility.Collapsed;
            SetCalendarStatus(MeetingOrganizerStatus, preview ? "Previewing change" : "Submitting organizer change", "", InfoBarSeverity.Informational);
            var submitted = false;
            try
            {
                request = await _exo.ValidateMeetingOrganizerChangeAsync(request);
                submitted = true;
                var matches = preview ? await _exo.PreviewMeetingOrganizerChangeAsync(request) : await _exo.ChangeMeetingOrganizerAsync(request);
                _exo.OperationCancellationToken.ThrowIfCancellationRequested();
                if (matches.Count > 0)
                {
                    foreach (var match in matches)
                    {
                        _meetingMatches.Add(match);
                        RecordCalendarTransfer(request with { UseEventId = true, Selector = match.EventId }, "Selection required", match.Subject);
                    }
                    MeetingMatchesPanel.Visibility = Visibility.Visible;
                    SetCalendarStatus(MeetingOrganizerStatus, "Meeting selection required", "Exchange returned candidate meetings. No transfer is confirmed.", InfoBarSeverity.Warning);
                    Logger.Log($"Organizer change returned {matches.Count} candidate meeting(s); no transfer is confirmed.", LogType.Warning);
                }
                else if (preview)
                {
                    _validatedMeetingOrganizer = request;
                    var details = $"{request.CurrentOrganizer} to {request.NewOrganizer}. Exchange WhatIf completed without error; no changes made. This does not guarantee a subsequent transfer will succeed.";
                    RecordCalendarTransfer(request, "Preview (no changes)", details);
                    SetCalendarStatus(MeetingOrganizerStatus, "Preview completed", details, InfoBarSeverity.Informational);
                }
                else
                {
                    const string details = "Command completed; calendar outcome unconfirmed. Check both calendars and Teams details before retrying.";
                    RecordCalendarTransfer(request, "Unconfirmed", details);
                    SetCalendarStatus(MeetingOrganizerStatus, "Command completed", details, InfoBarSeverity.Warning);
                    Logger.WriteSummary("Change Meeting Organizer", new Dictionary<string, object>
                    {
                        [request.CurrentOrganizer] = new[] { $"New organizer: {request.NewOrganizer}", $"Meeting: {request.Selector}", details }
                    });
                }
            }
            catch (Exception exception)
            {
                var status = preview ? "Preview failed (no changes)" : submitted ? "Unconfirmed" : "Validation failed (not submitted)";
                if (exception is OperationCanceledException) status = preview ? "Preview cancelled (no changes)" : "Cancelled; outcome unconfirmed";
                RecordCalendarTransfer(request, status, exception.Message);
                throw;
            }
        }, preview);

        private void ListViewMeetingMatches_SelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            if (ListViewMeetingMatches.SelectedItem is not MeetingOrganizerMatch match) return;
            ComboBoxMeetingSelector.SelectedIndex = 0;
            TextBoxMeetingSelector.Text = match.EventId;
        }

        private async void ButtonExportCalendarTransfers_Click(object sender, RoutedEventArgs args) => await RunUiOperationAsync(async () =>
        {
            if (_calendarTransfers.Count == 0) return;
            var picker = new FileSavePicker
            {
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
                SuggestedFileName = $"EXOKit_MeetingOrganizer_{DateTime.Now:yyyyMMdd_HHmmss}"
            };
            picker.FileTypeChoices.Add("CSV", new List<string> { ".csv" });
            InitializeWithWindow.Initialize(picker, GetWindowHandle());
            var file = await picker.PickSaveFileAsync();
            if (file != null)
            {
                await FileIO.WriteLinesAsync(file, CalendarTransferRecord.CreateCsvLines(_calendarTransfers.ToArray()));
                Logger.Log($"Exported {_calendarTransfers.Count} organizer result(s) to {file.Path}.");
            }
        }, readOnly: true);

        private void ButtonClearCalendarTransfers_Click(object sender, RoutedEventArgs args)
        {
            _calendarTransfers.Clear();
            UpdateCalendarControls();
        }
    }
}