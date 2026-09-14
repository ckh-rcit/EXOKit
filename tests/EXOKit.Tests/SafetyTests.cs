using EXOKit.Services;
using Xunit;

namespace EXOKit.Tests
{
    public sealed class SafetyTests : IDisposable
    {
        private readonly string _directory = Path.Combine(Path.GetTempPath(), "EXOKitTests", Guid.NewGuid().ToString("N"));
        private const string Tenant = "11111111-1111-1111-1111-111111111111";
        private static Task NoDelay() => Task.CompletedTask;
        private static Exception ServerBug() => new InvalidOperationException("Write-ErrorMessage: Object reference not set to an instance of an object");
        private static System.Collections.ObjectModel.Collection<System.Management.Automation.Host.ChoiceDescription> PublisherChoices() => new()
        {
            new("&Never run"), new("&Do not run"), new("&Run once"), new("&Always run")
        };

        public SafetyTests() => Directory.CreateDirectory(_directory);
        public void Dispose() => Directory.Delete(_directory, true);

        [Fact]
        public void HostedRunspaceCanImportLocalScriptModule()
        {
            var modulePath = Path.Combine(_directory, "PolicyProbe.psm1");
            File.WriteAllText(modulePath, "function Get-EXOKitPolicyProbe { 'module-loaded' }");
            var state = LoggerPSHost.CreateInitialSessionState();
            Assert.Equal(Microsoft.PowerShell.ExecutionPolicy.RemoteSigned, state.ExecutionPolicy);
            Assert.NotNull(state.AuthorizationManager);
            using var runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(new LoggerPSHost(), state);
            runspace.Open();
            using var pipeline = System.Management.Automation.PowerShell.Create();
            pipeline.Runspace = runspace;
            pipeline.AddCommand("Import-Module").AddParameter("Name", modulePath).AddParameter("ErrorAction", "Stop");
            pipeline.Invoke();
            Assert.False(pipeline.HadErrors);
            pipeline.Commands.Clear();
            pipeline.AddCommand("Get-EXOKitPolicyProbe");
            Assert.Equal("module-loaded", Assert.Single(pipeline.Invoke()).ToString());
        }

        [Fact]
        public void HostedRunspaceRejectsUnsignedInternetModule()
        {
            var modulePath = Path.Combine(_directory, "InternetPolicyProbe.psm1");
            File.WriteAllText(modulePath, "function Get-EXOKitInternetProbe { 'must-not-run' }");
            File.WriteAllText(modulePath + ":Zone.Identifier", "[ZoneTransfer]\r\nZoneId=3\r\n");
            using var runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(
                new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
            runspace.Open();
            using var pipeline = System.Management.Automation.PowerShell.Create();
            pipeline.Runspace = runspace;
            pipeline.AddCommand("Import-Module").AddParameter("Name", modulePath).AddParameter("ErrorAction", "Stop");
            var exception = Assert.Throws<System.Management.Automation.CmdletInvocationException>(() => pipeline.Invoke());
            Assert.Contains("not digitally signed", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PublisherPromptWithoutUiDoesNotSilentlyChooseDefault()
        {
            var choices = PublisherChoices();
            var host = new LoggerPSHost();
            Assert.Throws<InvalidOperationException>(() => host.UI.PromptForChoice(
                "Do you want to run software from this untrusted publisher?", "PackageManagement.format.ps1xml", choices, 1));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        public void PublisherPromptReturnsOnlyExplicitUserChoice(int selectedChoice)
        {
            var choices = PublisherChoices();
            var calls = 0;
            var host = new LoggerPSHost((caption, message, offeredChoices, defaultChoice) =>
            {
                calls++;
                Assert.Equal("Publisher confirmation", caption);
                Assert.Equal("PackageManagement.format.ps1xml", message);
                Assert.Same(choices, offeredChoices);
                Assert.Equal(1, defaultChoice);
                return selectedChoice;
            });
            Assert.Equal(selectedChoice, host.UI.PromptForChoice("Publisher confirmation", "PackageManagement.format.ps1xml", choices, 1));
            Assert.Equal(1, calls);
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(4)]
        public void PublisherPromptRejectsInvalidSelection(int selectedChoice)
        {
            var host = new LoggerPSHost((caption, message, choices, defaultChoice) => selectedChoice);
            Assert.Throws<InvalidOperationException>(() => host.UI.PromptForChoice("Publisher", "File", PublisherChoices(), 1));
        }

        [Fact]
        public void PublisherPromptCancellationDoesNotReturnDefault()
        {
            var host = new LoggerPSHost((caption, message, choices, defaultChoice) => throw new OperationCanceledException());
            Assert.Throws<OperationCanceledException>(() => host.UI.PromptForChoice("Publisher", "File", PublisherChoices(), 1));
        }

        [Fact]
        public async Task UnknownInitialStateNeverMutates()
        {
            var mutations = 0;
            var reads = 0;
            await Assert.ThrowsAsync<InvalidOperationException>(() => PermissionVerification.ApplyAsync(
                () => { mutations++; return Task.CompletedTask; },
                () => { reads++; throw ServerBug(); }, false, NoDelay));
            Assert.Equal(0, mutations);
            Assert.Equal(3, reads);
        }

        [Fact]
        public async Task ServerErrorAfterAppliedWriteIsVerifiedWithoutReplay()
        {
            var present = false;
            var mutations = 0;
            await PermissionVerification.ApplyAsync(
                () => { mutations++; present = true; throw ServerBug(); },
                () => Task.FromResult(present), true, NoDelay);
            Assert.Equal(1, mutations);
        }

        [Fact]
        public async Task SuccessfulButUnconfirmedWriteIsNotReplayed()
        {
            var mutations = 0;
            var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => PermissionVerification.ApplyAsync(
                () => { mutations++; return Task.CompletedTask; },
                () => Task.FromResult(false), true, NoDelay));
            Assert.Contains("unconfirmed", exception.Message);
            Assert.Equal(1, mutations);
        }

        [Fact]
        public async Task PersistentServerErrorIsBoundedAndUnconfirmed()
        {
            var mutations = 0;
            await Assert.ThrowsAsync<InvalidOperationException>(() => PermissionVerification.ApplyAsync(
                () => { mutations++; throw ServerBug(); }, () => Task.FromResult(false), true, NoDelay));
            Assert.Equal(2, mutations);
        }

        [Fact]
        public async Task PermissionErrorsDoNotBecomeAbsence()
        {
            var reads = 0;
            await Assert.ThrowsAsync<UnauthorizedAccessException>(() => PermissionVerification.ReadAsync<bool>(
                () => { reads++; throw new UnauthorizedAccessException(); }, NoDelay));
            Assert.Equal(1, reads);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task AlreadyDesiredStateDoesNotMutate(bool expected)
        {
            await PermissionVerification.ApplyAsync(() => throw new Exception("Unexpected mutation"), () => Task.FromResult(expected), expected, NoDelay);
        }

        [Fact]
        public async Task CancelledReadDoesNotRetry()
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() => PermissionVerification.ReadAsync<bool>(() => throw new OperationCanceledException(), NoDelay));
        }

        [Fact]
        public void SnapshotIsAtomicAndTenantBound()
        {
            var service = new SnapshotService(_directory) { TenantIdProvider = () => Tenant };
            var record = new SnapshotRecord
            {
                OperationType = "GroupDelegateRemoval", Target = "group@example.org",
                Items = { new SnapshotItem { User = "user@example.org", Role = "Send As" } }
            };
            var path = service.SaveSnapshot(record);
            var restored = service.LoadSnapshot(path)!;
            Assert.Equal(Tenant, restored.Metadata["TenantId"]);
            Assert.Single(restored.Items);
            Assert.Empty(Directory.GetFiles(service.SnapshotsDirectory, "*.tmp"));
        }

        [Fact]
        public void SnapshotWithoutTenantCannotBeSaved()
        {
            var service = new SnapshotService(_directory);
            Assert.Throws<InvalidOperationException>(() => service.SaveSnapshot(new SnapshotRecord()));
            Assert.False(Directory.Exists(service.SnapshotsDirectory));
        }

        [Fact]
        public void InvalidConfigurationDoesNotReplaceWorkingFile()
        {
            var config = ConfigService.CreateDefault();
            ConfigService.Save(config, _directory);
            var path = Path.Combine(_directory, "config.json");
            var previous = File.ReadAllText(path);
            config.Settings.GraphApi.Scopes.Clear();
            Assert.Throws<InvalidOperationException>(() => ConfigService.Save(config, _directory));
            Assert.Equal(previous, File.ReadAllText(path));
        }

        [Theory]
        [InlineData("GroupDelegateRemoval", "Send As")]
        [InlineData("GroupDelegateRemoval", "Send on Behalf")]
        [InlineData("MailboxPermissionRemoval", "Full Access")]
        [InlineData("GroupMembershipRemoval", "Owner")]
        public void SupportedRecoveryRolesValidate(string operation, string role)
        {
            var record = RecoveryRecord(operation, role);
            SnapshotService.ValidateForRestore(record, Tenant);
        }

        [Theory]
        [InlineData("Unknown", "Send As")]
        [InlineData("GroupDelegateRemoval", "Full Access")]
        [InlineData("GroupMembershipRemoval", "Unknown")]
        public void UnsupportedRecoveryCannotReachMutation(string operation, string role)
        {
            Assert.Throws<InvalidOperationException>(() => SnapshotService.ValidateForRestore(RecoveryRecord(operation, role), Tenant));
        }

        [Fact]
        public void LegacyOrCrossTenantRecoveryIsRejected()
        {
            var record = RecoveryRecord("GroupDelegateRemoval", "Send As");
            Assert.Throws<InvalidOperationException>(() => SnapshotService.ValidateForRestore(record, Guid.NewGuid().ToString()));
            record.Metadata.Clear();
            Assert.Throws<InvalidOperationException>(() => SnapshotService.ValidateForRestore(record, Tenant));
        }

        [Fact]
        public void EmptyRecoveryIsRejected()
        {
            var record = RecoveryRecord("GroupDelegateRemoval", "Send As");
            record.Items.Clear();
            Assert.Throws<InvalidOperationException>(() => SnapshotService.ValidateForRestore(record, Tenant));
        }

        private static SnapshotRecord RecoveryRecord(string operation, string role) => new()
        {
            OperationType = operation, Target = "group@example.org",
            Metadata = { ["TenantId"] = Tenant },
            Items = { new SnapshotItem { User = "user@example.org", Role = role } }
        };

        [Fact]
        public void ConfigurationSaveBacksUpPreviousFile()
        {
            var path = Path.Combine(_directory, "config.json");
            File.WriteAllText(path, "invalid prior configuration");
            ConfigService.Save(ConfigService.CreateDefault(), _directory);
            Assert.Equal("invalid prior configuration", File.ReadAllText(path + ".bak"));
            Assert.NotNull(ConfigService.Load(_directory));
        }

        [Theory]
        [InlineData("http://example.service-now.com/")]
        [InlineData("https://user:password@example.service-now.com/")]
        [InlineData("https://example.service-now.com/?redirect=bad")]
        [InlineData("https://example.service-now.com/other")]
        public void ServiceNowRejectsUnsafeEndpoint(string endpoint)
        {
            var config = SafeServiceNowConfig();
            config.InstanceUrl = endpoint;
            Assert.Throws<InvalidOperationException>(() => ServiceNowService.ValidateConfiguration(config));
        }

        [Fact]
        public void ServiceNowRejectsQueryInjectionInField()
        {
            var config = SafeServiceNowConfig();
            config.TicketNumberField = "number^ORactive";
            Assert.Throws<InvalidOperationException>(() => ServiceNowService.ValidateConfiguration(config));
        }

        [Fact]
        public async Task ServiceNowRejectsInjectedTicketBeforeCredentials()
        {
            var service = new ServiceNowService(SafeServiceNowConfig(), new GraphApiConfig());
            var result = await service.CloseTaskAsync("TASK1^ORactive=true");
            Assert.False(result.Success);
            Assert.Equal("Invalid ticket number.", result.Message);
        }

        private static ServiceNowConfig SafeServiceNowConfig() => new()
        {
            Enabled = true, InstanceUrl = "https://example.service-now.com/", KeyVaultUrl = "https://example.vault.azure.net/"
        };
    }
}

namespace EXOKit.Services
{
    public enum LogType { Info, Warning, Error }
    public static class Logger
    {
        public static void Log(string message, LogType type = LogType.Info) { }
    }
}