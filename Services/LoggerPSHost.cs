using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Text;

namespace EXOKit.Services
{
    /// <summary>
    /// Minimal PSHost implementation that forwards output to Logger and choice prompts to the UI.
    /// Security decisions require an explicit response; unavailable UI never implies consent.
    /// Ported from Entra Scout's LoggerPSHost.
    /// </summary>
    public class LoggerPSHost : PSHost
    {
        private readonly LoggerPSHostUserInterface _ui;
        private readonly Guid _instanceId = Guid.NewGuid();

        public LoggerPSHost(Func<string, string, Collection<ChoiceDescription>, int, int>? promptForChoice = null)
        {
            _ui = new LoggerPSHostUserInterface(promptForChoice);
        }

        public override CultureInfo CurrentCulture => CultureInfo.CurrentCulture;
        public override CultureInfo CurrentUICulture => CultureInfo.CurrentUICulture;
        public override Guid InstanceId => _instanceId;
        public override string Name => "EXOKitHost";
        public override PSHostUserInterface UI => _ui;
        public override Version Version => new(1, 0, 0);

        public override void EnterNestedPrompt() { }
        public override void ExitNestedPrompt() { }
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
        public override void SetShouldExit(int exitCode) { }
    }

    public class LoggerPSHostUserInterface : PSHostUserInterface
    {
        private readonly LoggerPSHostRawUserInterface _rawUi = new();
        private readonly StringBuilder _lineBuffer = new();
        private readonly Func<string, string, Collection<ChoiceDescription>, int, int>? _promptForChoice;

        public LoggerPSHostUserInterface(Func<string, string, Collection<ChoiceDescription>, int, int>? promptForChoice = null)
        {
            _promptForChoice = promptForChoice;
        }

        public override PSHostRawUserInterface RawUI => _rawUi;

        public override string ReadLine() => string.Empty;

        public override System.Security.SecureString ReadLineAsSecureString() => new();

        public override void Write(string value) => _lineBuffer.Append(value);

        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) => Write(value);

        public override void WriteLine(string value)
        {
            FlushLineBuffer();
            Logger.Log(value);
        }

        public override void WriteLine()
        {
            FlushLineBuffer();
        }

        public override void WriteErrorLine(string value)
        {
            FlushLineBuffer();
            Logger.Log($"ERROR: {value}", LogType.Error);
        }

        public override void WriteDebugLine(string message) => Logger.Log($"DEBUG: {message}");

        public override void WriteProgress(long sourceId, System.Management.Automation.ProgressRecord record) { }

        public override void WriteVerboseLine(string message) => Logger.Log($"VERBOSE: {message}");

        public override void WriteWarningLine(string message) => Logger.Log($"WARNING: {message}", LogType.Warning);

        // Note: warnings are still logged as LogType.Warning here so they're captured in Logger's
        // history; the UI (MainWindow) is responsible for hiding LogType.Warning lines by default
        // and revealing them when the user checks "Show Warnings".

        public override System.Collections.Generic.Dictionary<string, System.Management.Automation.PSObject> Prompt(
            string caption, string message, System.Collections.ObjectModel.Collection<System.Management.Automation.Host.FieldDescription> descriptions)
        {
            Logger.Log($"{caption}: {message}");
            return new System.Collections.Generic.Dictionary<string, System.Management.Automation.PSObject>();
        }

        public override int PromptForChoice(string caption, string message, System.Collections.ObjectModel.Collection<ChoiceDescription> choices, int defaultChoice)
        {
            Logger.Log($"{caption}: {message}");
            if (_promptForChoice == null)
                throw new InvalidOperationException("PowerShell requires an explicit choice, but no prompt UI is available. No choice was made.");
            var selectedChoice = _promptForChoice(caption, message, choices, defaultChoice);
            if (selectedChoice < 0 || selectedChoice >= choices.Count)
                throw new InvalidOperationException("The PowerShell prompt did not return a valid choice. No choice was made.");
            return selectedChoice;
        }

        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName) =>
            PromptForCredential(caption, message, userName, targetName, PSCredentialTypes.Default, PSCredentialUIOptions.Default);

        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName,
            PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
        {
            Logger.Log($"{caption}: {message}");
            return new PSCredential(userName ?? string.Empty, new System.Security.SecureString());
        }

        private void FlushLineBuffer()
        {
            if (_lineBuffer.Length == 0)
            {
                return;
            }

            Logger.Log(_lineBuffer.ToString());
            _lineBuffer.Clear();
        }
    }

    public class LoggerPSHostRawUserInterface : PSHostRawUserInterface
    {
        public override ConsoleColor ForegroundColor { get; set; } = ConsoleColor.Gray;
        public override ConsoleColor BackgroundColor { get; set; } = ConsoleColor.Black;
        public override Coordinates CursorPosition { get; set; } = new(0, 0);
        public override Coordinates WindowPosition { get; set; } = new(0, 0);
        public override int CursorSize { get; set; } = 25;
        public override Size BufferSize { get; set; } = new(120, 9999);
        public override Size WindowSize { get; set; } = new(120, 40);
        public override Size MaxWindowSize => new(120, 40);
        public override Size MaxPhysicalWindowSize => new(120, 40);
        public override string WindowTitle { get; set; } = "EXOKit";
        public override bool KeyAvailable => false;

        public override void FlushInputBuffer() { }

        public override BufferCell[,] GetBufferContents(Rectangle rectangle) => new BufferCell[0, 0];

        public override KeyInfo ReadKey(ReadKeyOptions options) => new();

        public override void ScrollBufferContents(Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill) { }

        public override void SetBufferContents(Rectangle rectangle, BufferCell fill) { }

        public override void SetBufferContents(Coordinates origin, BufferCell[,] contents) { }
    }
}
