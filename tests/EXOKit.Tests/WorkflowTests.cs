using EXOKit.Services;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Xunit;

namespace EXOKit.Tests;

public sealed class WorkflowTests
{
    [Fact]
    public void HostedRuntimeLoadsBundledCoreCommandsWithOnlyDesktopModulePaths()
    {
        using var runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            $savedModulePath = $env:PSModulePath
            try {
                $env:PSModulePath = [IO.Path]::Combine($env:WINDIR, 'System32', 'WindowsPowerShell', 'v1.0', 'Modules')
                $ErrorActionPreference = 'Stop'
                Join-Path 'C:\' 'Modules'
                [pscustomobject]@{ Ready = $true } | ConvertTo-Json -Compress
                (Get-Module Microsoft.PowerShell.Management).ModuleBase
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var results = powershell.Invoke();
        Assert.False(powershell.HadErrors, string.Join(Environment.NewLine, powershell.Streams.Error));
        Assert.Equal("C:\\Modules", results[0].ToString());
        Assert.Equal("{\"Ready\":true}", results[1].ToString());
        Assert.StartsWith(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar), results[2].ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EmbeddedHostCanInstallAndImportExoPrerequisitesWithoutAuthentication()
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $moduleRoot)
            $savedModulePath = $env:PSModulePath
            $env:PSModulePath = [IO.Path]::Combine($PSHOME, 'Modules')
            try {
                & ([scriptblock]::Create($bootstrap)) -ModuleRoot $moduleRoot
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var moduleRoot = Path.Combine(Path.GetTempPath(), "EXOKit-bootstrap-" + Guid.NewGuid().ToString("N"));
        powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("moduleRoot", moduleRoot);
        powershell.Invoke();
        Assert.False(powershell.HadErrors, string.Join(Environment.NewLine, powershell.Streams.Error));
        powershell.Commands.Clear();
        powershell.AddCommand("Get-Command").AddParameter("Name", "Connect-ExchangeOnline").AddParameter("ErrorAction", "Stop");
        var commands = powershell.Invoke();
        Assert.False(powershell.HadErrors, string.Join(Environment.NewLine, powershell.Streams.Error));
        var command = (CommandInfo)Assert.Single(commands).BaseObject;
        Assert.StartsWith(moduleRoot, command.Module.ModuleBase, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("Connected", "Active", true)]
    [InlineData("Connected", "Expired", false)]
    [InlineData("Disconnected", "Active", false)]
    public void AuthenticationRequiresActiveExchangeToken(string state, string token, bool valid)
    {
        var connection = new PSObject();
        connection.Properties.Add(new PSNoteProperty("State", state));
        connection.Properties.Add(new PSNoteProperty("TokenStatus", token));
        if (valid) ExoPowerShellService.ValidateConnection(connection);
        else Assert.Throws<InvalidOperationException>(() => ExoPowerShellService.ValidateConnection(connection));
    }

    [Fact]
    public void AuthenticationRetriesWamOnlyThroughBrowser()
    {
        using var runspace = CreateRunspace("""
            function Connect-ExchangeOnline { [CmdletBinding()] param([switch]$DisableWAM,[switch]$SkipLoadingFormatData,$ShowBanner)
                $global:attempts++
                if (!$DisableWAM) { throw 'A window handle must be configured.' }
            }
            """);
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        ExoPowerShellService.InvokeInteractiveConnection(powershell, false);
        Assert.Equal(2, runspace.SessionStateProxy.GetVariable("attempts"));
    }

    [Fact]
    public void AuthenticationDoesNotRetryConsentFailure()
    {
        using var runspace = CreateRunspace("""
            function Connect-ExchangeOnline { [CmdletBinding()] param([switch]$DisableWAM,[switch]$SkipLoadingFormatData,$ShowBanner)
                $global:attempts++
                throw 'AADSTS65001 consent required'
            }
            """);
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        Assert.ThrowsAny<Exception>(() => ExoPowerShellService.InvokeInteractiveConnection(powershell, false));
        Assert.Equal(1, runspace.SessionStateProxy.GetVariable("attempts"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExchangePipelineCancellationAndTimeoutStopActiveWork(bool timeout)
    {
        using var runspace = CreateRunspace("""
            function Invoke-BlockedQuery { [CmdletBinding()] param()
                [void]$global:QueryStarted.TrySetResult($true)
                while ($true) { [Threading.Thread]::SpinWait(1000) }
            }
            """);
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        runspace.SessionStateProxy.SetVariable("QueryStarted", started);
        using var cancellation = new CancellationTokenSource();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddCommand("Invoke-BlockedQuery");
        var query = Task.Run(() => ExoPowerShellService.InvokePipelineCore(powershell, cancellation.Token,
            timeout ? TimeSpan.FromMilliseconds(500) : TimeSpan.FromSeconds(10)));
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        if (timeout)
            Assert.Contains("timed out", (await Assert.ThrowsAsync<TimeoutException>(async () => await query.WaitAsync(TimeSpan.FromSeconds(5)))).Message);
        else
        {
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await query.WaitAsync(TimeSpan.FromSeconds(5)));
        }
        powershell.Commands.Clear();
        powershell.AddScript("'runspace reusable'");
        Assert.Equal("runspace reusable", Assert.Single(ExoPowerShellService.InvokePipelineCore(powershell, default, TimeSpan.FromSeconds(5))).ToString());
    }

    [Fact]
    public async Task MailboxBatchPreservesOtherResultsAfterLookupFailure()
    {
        using var runspace = CreateRunspace("""
            function Get-Mailbox { [CmdletBinding()] param($Identity) [pscustomobject]@{Identity=$Identity;GrantSendOnBehalfTo=[Collections.ArrayList]@()} }
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                if ($Identity -eq 'bad@example.org') { throw 'Access denied' }
                [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails='UserMailbox'}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var results = await new MailboxPermissionService(exo).InvokePermissionOperationAsync(PermissionOperationType.Remove, PermissionTargetType.Mailbox,
                new[] { "mailbox" }, new[] { "bad@example.org", "good@example.org" }, new PermissionSelections { SendOnBehalf = true });
            Assert.Equal(2, results.Count);
            Assert.Contains("User Lookup Error", results.Single(result => result.User == "bad@example.org").Statuses);
            Assert.Contains("Send on Behalf (Not Found - No Change)", results.Single(result => result.User == "good@example.org").Statuses);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task PermissionReportsPreserveArrayListGrantsAndUnresolvedSids()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                [pscustomobject]@{Identity='mailbox';PrimarySmtpAddress='mailbox@example.org';RecipientTypeDetails='SharedMailbox'}
            }
            function Get-EXOMailboxPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                if ($PrimarySmtpAddress -ne 'mailbox@example.org') { throw 'Permission query must use resolved primary SMTP address' }
                [pscustomobject]@{User='user@example.org';AccessRights=[Collections.ArrayList]@('FullAccess');Deny=$false;IsInherited=$false}
                [pscustomobject]@{User='S-1-5-21-999';AccessRights=[Collections.ArrayList]@('FullAccess');Deny=$false;IsInherited=$false}
                [pscustomobject]@{User='denied';AccessRights=[Collections.ArrayList]@('FullAccess');Deny=$true;IsInherited=$false}
            }
            function Get-EXORecipientPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                if ($PrimarySmtpAddress -ne 'mailbox@example.org') { throw 'Permission query must use resolved primary SMTP address' }
                [pscustomobject]@{Trustee='user@example.org';AccessRights=[Collections.ArrayList]@('SendAs');AccessControlType='Allow';IsInherited=$false}
                [pscustomobject]@{Trustee='inherited';AccessRights=[Collections.ArrayList]@('SendAs');AccessControlType='Allow';IsInherited=$true}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            Assert.Equal(new[] { "user@example.org", "S-1-5-21-999" }, await exo.GetAllFullAccessDelegatesAsync("mailbox"));
            Assert.Equal(new[] { "user@example.org" }, await exo.GetAllSendAsDelegatesAsync("mailbox"));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public void ReportCsvIncludesWarningDetails()
    {
        var row = new ReportRow
        {
            ObjectName = "mailbox", MemberOrDelegate = "deleted-delegate", RoleOrPermission = "Send on Behalf",
            Status = "Unresolved: recipient \"deleted\", unavailable"
        };
        var lines = ReportingService.CreateCsvLines(new[] { row }).ToArray();
        Assert.EndsWith(",Status", lines[0]);
        Assert.Contains("deleted-delegate", lines[1]);
        Assert.EndsWith(InputParsingHelpers.EscapeCsv(row.Status), lines[1]);
        Assert.True(row.HasWarning);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PermissionReportsRejectOutOfScopeRows(bool sendAs)
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                [pscustomobject]@{Identity='target';PrimarySmtpAddress='target@example.org';RecipientTypeDetails='SharedMailbox'}
            }
            function Get-EXOMailboxPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                [pscustomobject]@{Identity='other@example.org';User='delegate@example.org';AccessRights=[Collections.ArrayList]@('FullAccess')}
            }
            function Get-EXORecipientPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                [pscustomobject]@{Identity='other@example.org';Trustee='delegate@example.org';AccessRights=[Collections.ArrayList]@('SendAs')}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var error = await Assert.ThrowsAsync<InvalidOperationException>(() => sendAs
                ? exo.GetAllSendAsDelegatesAsync("target") : exo.GetAllFullAccessDelegatesAsync("target"));
            Assert.Contains("no out-of-scope rows were included", error.Message);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReportCancellationCannotReturnSuccessAndAllowsAnotherQuery(bool cancelBeforeStart)
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                if ($global:BlockQuery) {
                    [void]$global:QueryStarted.TrySetResult($true)
                    while ($true) { [Threading.Thread]::SpinWait(1000) }
                }
                [pscustomobject]@{PrimarySmtpAddress='target@example.org';RecipientTypeDetails='SharedMailbox'}
            }
            function Get-EXOMailboxPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                [pscustomobject]@{Identity=$PrimarySmtpAddress;User='delegate@example.org';AccessRights=[Collections.ArrayList]@('FullAccess')}
            }
            """);
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        runspace.SessionStateProxy.SetVariable("QueryStarted", started);
        runspace.SessionStateProxy.SetVariable("BlockQuery", true);
        using var cancellation = new CancellationTokenSource();
        var exo = new ExoPowerShellService(runspace) { OperationCancellationToken = cancellation.Token };
        try
        {
            if (cancelBeforeStart) cancellation.Cancel();
            var query = exo.GetAllFullAccessDelegatesAsync("target");
            if (!cancelBeforeStart)
            {
                await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
                cancellation.Cancel();
            }
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await query.WaitAsync(TimeSpan.FromSeconds(5)));
            if (cancelBeforeStart) Assert.False(started.Task.IsCompleted);
            exo.OperationCancellationToken = default;
            runspace.SessionStateProxy.SetVariable("BlockQuery", false);
            Assert.Single(await exo.GetAllFullAccessDelegatesAsync("target"));
        }
        finally { exo.OperationCancellationToken = default; await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MailboxReportPreservesReadableSectionsAndUnresolvedDelegates(bool denySendAs)
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                if ($Identity -eq 'deleted-delegate') { return }
                [pscustomobject]@{Identity=$Identity;Name=$Identity;PrimarySmtpAddress=$Identity;RecipientTypeDetails='SharedMailbox'}
            }
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                [pscustomobject]@{GrantSendOnBehalfTo=[Collections.ArrayList]@('deleted-delegate','known@example.org')}
            }
            function Get-EXOMailboxPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                [pscustomobject]@{User='full@example.org';AccessRights=[Collections.ArrayList]@('FullAccess')}
            }
            function Get-EXORecipientPermission { [CmdletBinding()] param($PrimarySmtpAddress,$ResultSize)
                if ($global:DenySendAs) { throw 'Send As access denied' }
                [pscustomobject]@{Trustee='send@example.org';AccessRights=[Collections.ArrayList]@('SendAs')}
            }
            """);
        runspace.SessionStateProxy.SetVariable("DenySendAs", denySendAs);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var rows = await new ReportingService(exo).GetMailboxDelegatesAsync("mailbox@example.org");
            Assert.Contains(rows, row => row.MemberOrDelegate == "full@example.org" && !row.HasWarning);
            Assert.Contains(rows, row => row.MemberOrDelegate == "known@example.org" && !row.HasWarning);
            Assert.Contains(rows, row => row.MemberOrDelegate == "deleted-delegate" && row.HasWarning);
            Assert.Equal(denySendAs, rows.Single(row => row.RoleOrPermission == "Send As").HasWarning);
            await Assert.ThrowsAsync<InvalidOperationException>(() => exo.GetAllSendOnBehalfDelegatesAsync("mailbox@example.org"));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task PermissionReportsRejectUnresolvedTargetsBeforeQuery()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails) }
            function Get-EXOMailboxPermission { throw 'Unscoped query must not execute' }
            function Get-EXORecipientPermission { throw 'Unscoped query must not execute' }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            Assert.Contains("No permission query was run", (await Assert.ThrowsAsync<InvalidOperationException>(() => exo.GetAllFullAccessDelegatesAsync("missing"))).Message);
            Assert.Contains("No permission query was run", (await Assert.ThrowsAsync<InvalidOperationException>(() => exo.GetAllSendAsDelegatesAsync("missing"))).Message);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task DistributionGroupReportPreservesMembersWhenOwnerIsUnresolved()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                if ($Identity -eq 'deleted-owner') { return }
                [pscustomobject]@{Identity='group';PrimarySmtpAddress='group@example.org';RecipientTypeDetails='MailUniversalDistributionGroup'}
            }
            function Get-DistributionGroup { [CmdletBinding()] param($Identity)
                [pscustomobject]@{ManagedBy=[Collections.ArrayList]@('deleted-owner')}
            }
            function Get-DistributionGroupMember { [CmdletBinding()] param($Identity,$ResultSize)
                [pscustomobject]@{DisplayName='Member';PrimarySmtpAddress='member@example.org'}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var rows = await new ReportingService(exo).GenerateObjectReportAsync("group");
            Assert.Contains(rows, row => row.MemberOrDelegate == "deleted-owner" && row.HasWarning);
            Assert.Contains(rows, row => row.RoleOrPermission == "Member" && !row.HasWarning);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task DistributionGroupOwnerCheckResolvesSerializedIdentity()
    {
        using var runspace = CreateRunspace("""
            function Get-DistributionGroup { [CmdletBinding()] param($Identity)
                [pscustomobject]@{ManagedBy=[Collections.ArrayList]@('serialized-owner')}
            }
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                [pscustomobject]@{PrimarySmtpAddress='owner@example.org'}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try { Assert.True(await exo.IsDistributionGroupOwnerAsync("group", new RecipientInfo { PrimarySmtpAddress = "owner@example.org" }, "owner@example.org")); }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task DistributionGroupReportFailsWhenOwnerCannotBeRead()
    {
        using var runspace = CreateRunspace("""
            function Get-DistributionGroup { [CmdletBinding()] param($Identity)
                [pscustomobject]@{ManagedBy=[Collections.ArrayList]@('owner')}
            }
            function Get-Recipient { [CmdletBinding()] param($Identity)
                $PSCmdlet.WriteError([Management.Automation.ErrorRecord]::new([UnauthorizedAccessException]::new('Access denied'), 'AccessDenied', [Management.Automation.ErrorCategory]::PermissionDenied, $Identity))
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try { await Assert.ThrowsAsync<InvalidOperationException>(() => exo.GetDistributionGroupReportLinksAsync("group", false)); }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task RecipientResolverFallsBackOnlyForTypedMissingObject()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                if (!$RecipientTypeDetails) {
                    $PSCmdlet.WriteError([Management.Automation.ErrorRecord]::new([InvalidOperationException]::new('Recipient absent'), 'ManagementObjectNotFoundException', [Management.Automation.ErrorCategory]::ObjectNotFound, $Identity))
                } else {
                    [pscustomobject]@{RecipientTypeDetails='GroupMailbox';PrimarySmtpAddress=$Identity}
                }
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try { Assert.Equal("GroupMailbox", (await exo.GetRecipientAsync("group@example.org"))?.RecipientTypeDetails); }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData(true, true, true)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    public async Task BookingsPreflightRequiresOrganizationAndPolicy(bool organizationEnabled, bool policyEnabled, bool expected)
    {
        using var runspace = CreateRunspace("""
            function Get-OrganizationConfig { [CmdletBinding()] param() [pscustomobject]@{BookingsEnabled=$global:OrganizationEnabled} }
            function Get-OwaMailboxPolicy { [CmdletBinding()] param($Identity) [pscustomobject]@{BookingsMailboxCreationEnabled=$global:PolicyEnabled} }
            """);
        runspace.SessionStateProxy.SetVariable("OrganizationEnabled", organizationEnabled);
        runspace.SessionStateProxy.SetVariable("PolicyEnabled", policyEnabled);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            if (expected) await exo.ValidateBookingsConfigurationAsync("BookingsCreators");
            else await Assert.ThrowsAsync<InvalidOperationException>(() => exo.ValidateBookingsConfigurationAsync("BookingsCreators"));
        }
        finally { await exo.ShutdownAsync(); }
    }

    private sealed class CancelAfterLookupHandler(CancellationTokenSource cancellation) : HttpMessageHandler
    {
        public int Requests { get; private set; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests++;
            Assert.Equal(HttpMethod.Get, request.Method);
            cancellation.Cancel();
            return Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent("""{"result":[{"sys_id":"11111111111111111111111111111111","number":"TASK001","sys_class_name":"task"}]}""")
            });
        }
    }

    [Fact]
    public async Task CancellationAfterTicketLookupPreventsPatch()
    {
        using var cancellation = new CancellationTokenSource();
        using var handler = new CancelAfterLookupHandler(cancellation);
        using var client = new HttpClient(handler);
        var config = new ServiceNowConfig
        {
            Enabled = true, InstanceUrl = "https://example.service-now.com/", KeyVaultUrl = "https://example.vault.azure.net/",
            UsernameSecretName = "test-user", PasswordSecretName = "test-password"
        };
        var service = new ServiceNowService(config, client, (_, token) => Task.FromResult("synthetic-test-value"));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => service.CloseTaskAsync("TASK001", cancellationToken: cancellation.Token));
        Assert.Equal(1, handler.Requests);
    }

    [Fact]
    public async Task AbsentSendOnBehalfRemovalHasFinalNoChangeStatus()
    {
        using var runspace = CreateRunspace("""
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                [pscustomobject]@{Identity=$Identity;GrantSendOnBehalfTo=[Collections.ArrayList]@()}
            }
            function Get-Recipient { [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails='UserMailbox'}
            }
            function Set-Mailbox { throw 'An absent delegate must never be removed' }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var service = new MailboxPermissionService(exo);
            var results = await service.InvokePermissionOperationAsync(PermissionOperationType.Remove, PermissionTargetType.Mailbox,
                new[] { "mailbox@example.org" }, new[] { "user@example.org" }, new PermissionSelections { SendOnBehalf = true });
            Assert.Equal(new[] { "Send on Behalf (Not Found - No Change)" }, Assert.Single(results).Statuses);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task SnapshotFailureCannotBecomePermissionNotFound()
    {
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var blocked = Path.Combine(directory, "workspace");
        File.WriteAllText(blocked, "not a directory");
        using var runspace = CreateRunspace("""
            function Get-Mailbox { [CmdletBinding()] param($Identity) [pscustomobject]@{Identity=$Identity} }
            function Get-Recipient { [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails='UserMailbox'}
            }
            function Get-EXOMailboxPermission { [CmdletBinding()] param($Identity,$ResultSize)
                [pscustomobject]@{User='user@example.org';AccessRights=[Collections.ArrayList]@('FullAccess');Deny=$false}
            }
            function Remove-MailboxPermission { throw 'Snapshot failure must prevent mutation' }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var snapshots = new SnapshotService(blocked) { TenantIdProvider = () => "11111111-1111-1111-1111-111111111111" };
            var service = new MailboxPermissionService(exo, snapshots);
            var results = await service.InvokePermissionOperationAsync(PermissionOperationType.Remove, PermissionTargetType.Mailbox,
                new[] { "mailbox@example.org" }, new[] { "user@example.org" }, new PermissionSelections { FullAccess = true });
            Assert.Equal(new[] { "Full Access (Error)" }, Assert.Single(results).Statuses);
        }
        finally { await exo.ShutdownAsync(); Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task GroupSettingsSaveFailuresExposeActionableErrors()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity,$RecipientTypeDetails)
                [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails='MailUniversalDistributionGroup'}
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var service = new GroupSettingsService(exo);
            Assert.False(await service.SaveDeliveryManagementAsync("group", false, Array.Empty<string>()));
            Assert.Contains("Load the group's settings", service.LastError);
            Assert.False(await service.SaveDelegatesAsync("group", Array.Empty<string>(), Array.Empty<string>()));
            Assert.Contains("Load the group's settings", service.LastError);
            Assert.False(await service.SaveMessageApprovalAsync("group", false, Array.Empty<string>(), Array.Empty<string>(), "Always"));
            Assert.Contains("Load the group's settings", service.LastError);
            Assert.False(await service.SaveMembershipApprovalAsync("group", "Closed", "Closed"));
            Assert.Contains("Load the group's settings", service.LastError);
            Assert.False(await service.SaveMessageApprovalAsync("group", true, Array.Empty<string>(), Array.Empty<string>(), "Always"));
            Assert.Contains("moderator is required", service.LastError);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Fact]
    public async Task GroupSettingsRoundTripPreservesListsAndSnapshotsRemovals()
    {
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        using var runspace = CreateRunspace("""
            $global:Delegates=[Collections.ArrayList]@('first@example.org','second@example.org')
            $global:Senders=[Collections.ArrayList]@('sender@example.org')
            function Get-Recipient { [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails=$(if($Identity -eq 'group'){ 'MailUniversalDistributionGroup' }else{ 'UserMailbox' })}
            }
            function Get-DistributionGroup { [CmdletBinding()] param($Identity)
                [pscustomobject]@{GrantSendOnBehalfTo=$global:Delegates;AcceptMessagesOnlyFromSendersOrMembers=$global:Senders;
                    ModeratedBy=[Collections.ArrayList]@('moderator@example.org');BypassModerationFromSendersOrMembers=[Collections.ArrayList]@('bypass@example.org');
                    RequireSenderAuthenticationEnabled=$true;ModerationEnabled=$true;SendModerationNotifications='Always';MemberJoinRestriction='Closed';MemberDepartRestriction='Closed'}
            }
            function Get-RecipientPermission { [CmdletBinding()] param($Identity,$AccessRights) }
            function Set-DistributionGroup { [CmdletBinding(SupportsShouldProcess)] param($Identity,$GrantSendOnBehalfTo,$RequireSenderAuthenticationEnabled,$AcceptMessagesOnlyFromSendersOrMembers,[switch]$BypassSecurityGroupManagerCheck)
                if ($PSBoundParameters.ContainsKey('GrantSendOnBehalfTo')) { $global:Delegates=[Collections.ArrayList]@($GrantSendOnBehalfTo) }
                if ($PSBoundParameters.ContainsKey('AcceptMessagesOnlyFromSendersOrMembers')) { $global:Senders=[Collections.ArrayList]@($AcceptMessagesOnlyFromSendersOrMembers) }
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var snapshots = new SnapshotService(directory) { TenantIdProvider = () => "11111111-1111-1111-1111-111111111111" };
            var service = new GroupSettingsService(exo, snapshots);
            var loaded = await service.LoadSettingsAsync("group");
            Assert.Equal(2, loaded!.SendOnBehalfDelegates.Length);
            Assert.Single(loaded.Moderators);
            Assert.Single(loaded.BypassModerationSenders);
            runspace.SessionStateProxy.SetVariable("Delegates", new System.Collections.ArrayList { "second@example.org", "first@example.org" });
            Assert.True(await service.SaveDeliveryManagementAsync("group", false, loaded.SpecifiedSenders));
            Assert.Null(service.LastError);
            Assert.True(await service.SaveDelegatesAsync("group", loaded.SendAsDelegates, loaded.SendOnBehalfDelegates));
            Assert.Empty(snapshots.ListSnapshots());
            Assert.True(await service.SaveDelegatesAsync("group", Array.Empty<string>(), new[] { "first@example.org" }));
            Assert.Equal("second@example.org", Assert.Single(Assert.Single(snapshots.ListSnapshots()).Items).User);
            Assert.Equal(new[] { "first@example.org" }, (await service.LoadSettingsAsync("group"))!.SendOnBehalfDelegates);
        }
        finally { await exo.ShutdownAsync(); if (Directory.Exists(directory)) Directory.Delete(directory, true); }
    }

    private sealed class GraphHandler(string response) : HttpMessageHandler
    {
        public string? Request { get; private set; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Request = Uri.UnescapeDataString(request.RequestUri!.ToString());
            return Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent(response, System.Text.Encoding.UTF8, "application/json")
            });
        }
    }

    [Fact]
    public async Task GraphUserLookupUsesUniqueEmailAndAliasFilter()
    {
        using var handler = new GraphHandler("""{"value":[{"id":"user-id"}]}""");
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var auth = new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", "");
        var service = new GraphService(auth, client);
        Assert.Equal("user-id", await service.GetUserIdAsync("o'brien@example.org"));
        Assert.Contains("userPrincipalName eq 'o''brien@example.org'", handler.Request);
        Assert.Contains("mail eq 'o''brien@example.org'", handler.Request);
        Assert.Contains("proxyAddresses/any", handler.Request);
        auth.OperationCancellationToken = new CancellationToken(true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => service.GetUserIdAsync("user@example.org"));
    }

    [Theory]
    [InlineData("{\"value\":[{\"id\":\"one\"},{\"id\":\"two\"}]}")]
    [InlineData("{}")]
    [InlineData("{\"value\":[{}]}")]
    public async Task GraphUserLookupRejectsAmbiguousOrUnreadableResponses(string response)
    {
        using var handler = new GraphHandler(response);
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var service = new GraphService(new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", ""), client);
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetUserIdAsync("user@example.org"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task SendAsReadsEntireAclAndProtectsUnresolvedHistory(bool knownHistory)
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient {
                [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                if ($Identity -eq 'user@example.org') {
                    [pscustomobject]@{PrimarySmtpAddress=$Identity;RecipientTypeDetails='UserMailbox'}
                }
            }
            function Get-User {
                [CmdletBinding()] param($Identity)
                [pscustomobject]@{Sid='S-1-5-21-100';SidHistory=$global:History}
            }
            function Get-RecipientPermission {
                [CmdletBinding()] param($Identity, $ResultSize)
                if ($ResultSize -ne 'Unlimited') { throw 'Truncated ACL' }
                [pscustomobject]@{Trustee='S-1-5-21-99';AccessRights=[Collections.ArrayList]@('SendAs');AccessControlType='Allow'}
            }
            """);
        runspace.SessionStateProxy.SetVariable("History", knownHistory ? new[] { "S-1-5-21-99" } : Array.Empty<string>());
        var service = new ExoPowerShellService(runspace);
        try
        {
            if (knownHistory) Assert.True(await service.HasSendAsAsync("mailbox", "user@example.org"));
            else await Assert.ThrowsAsync<InvalidOperationException>(() => service.HasSendAsAsync("mailbox", "user@example.org"));
        }
        finally { await service.ShutdownAsync(); }
    }

    private static Runspace CreateRunspace(string script)
    {
        var runspace = RunspaceFactory.CreateRunspace(LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var pipeline = PowerShell.Create();
        pipeline.Runspace = runspace;
        pipeline.AddScript(script);
        pipeline.Invoke();
        Assert.False(pipeline.HadErrors);
        return runspace;
    }

    [Fact]
    public async Task RecipientResolverExplicitlyRequestsGroupMailbox()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient {
                [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                if ($RecipientTypeDetails -eq 'GroupMailbox') {
                    [pscustomobject]@{ PrimarySmtpAddress=$Identity; RecipientTypeDetails='GroupMailbox' }
                }
            }
            """);
        var service = new ExoPowerShellService(runspace);
        try
        {
            Assert.Equal("GroupMailbox", (await service.GetRecipientAsync("group@example.org"))?.RecipientTypeDetails);
            await Assert.ThrowsAsync<ArgumentException>(() => service.GetRecipientAsync(""));
        }
        finally { await service.ShutdownAsync(); }
    }

    [Fact]
    public async Task RecipientResolverDoesNotHideAccessDenied()
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient {
                [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                $PSCmdlet.WriteError([Management.Automation.ErrorRecord]::new([UnauthorizedAccessException]::new('Access denied'), 'AccessDenied', [Management.Automation.ErrorCategory]::PermissionDenied, $Identity))
            }
            """);
        var service = new ExoPowerShellService(runspace);
        try { await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetRecipientAsync("group@example.org")); }
        finally { await service.ShutdownAsync(); }
    }

    [Fact]
    public async Task DistributionGroupWritesUseAdministratorSwitch()
    {
        using var runspace = CreateRunspace("""
            function Add-DistributionGroupMember {
                [CmdletBinding(SupportsShouldProcess)] param($Identity, $Member, [switch]$BypassSecurityGroupManagerCheck)
                if (!$BypassSecurityGroupManagerCheck) { throw 'Manager check not enabled' }
            }
            function Remove-DistributionGroupMember {
                [CmdletBinding(SupportsShouldProcess)] param($Identity, $Member, [switch]$BypassSecurityGroupManagerCheck)
                if (!$BypassSecurityGroupManagerCheck) { throw 'Manager check not enabled' }
            }
            function Set-DistributionGroup {
                [CmdletBinding(SupportsShouldProcess)]
                param($Identity, $ManagedBy, $RequireSenderAuthenticationEnabled, $AcceptMessagesOnlyFromSendersOrMembers,
                      $GrantSendOnBehalfTo, $ModerationEnabled, $ModeratedBy, $BypassModerationFromSendersOrMembers,
                      $SendModerationNotifications, $MemberJoinRestriction, $MemberDepartRestriction,
                      [switch]$BypassSecurityGroupManagerCheck)
                if (!$BypassSecurityGroupManagerCheck) { throw 'Manager check not enabled' }
            }
            """);
        var service = new ExoPowerShellService(runspace);
        try
        {
            await service.AddDistributionGroupMemberAsync("group", "user");
            await service.RemoveDistributionGroupMemberAsync("group", "user");
            await service.AddDistributionGroupOwnerAsync("group", "user");
            await service.RemoveDistributionGroupOwnerAsync("group", "user");
            await service.SetDeliveryManagementAsync("group", false, new[] { "user" });
            await service.SetSendOnBehalfDelegatesAsync("group", new[] { "user" });
            await service.AddGroupSendOnBehalfAsync("group", "user");
            await service.SetMessageApprovalAsync("group", true, new[] { "user" }, Array.Empty<string>(), "Always");
            await service.SetMembershipApprovalAsync("group", "Closed", "Closed");
        }
        finally { await service.ShutdownAsync(); }
    }
}