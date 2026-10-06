using EXOKit.Services;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Xunit;

namespace EXOKit.Tests;

public sealed class WorkflowTests
{
    [Fact]
    public async Task CalendarOrganizerPreviewCannotMutateAndCsvPreservesUnconfirmedStatus()
    {
        using var runspace = CreateRunspace("""
            $global:Transfers = 0
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                [pscustomobject]@{ PrimarySmtpAddress = $Identity; RecipientTypeDetails = 'UserMailbox' }
            }
            function Invoke-ChangeMeetingOrganizer { [CmdletBinding(SupportsShouldProcess)] param($Identity, $NewOrganizer, $EventId)
                if (!$WhatIfPreference) { $global:Transfers++; throw 'Preview attempted a mutation' }
                $global:Previewed = $true
            }
            """);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var request = new MeetingOrganizerRequest("old@example.org", "new@example.org", true, "series-id", null);
            Assert.Empty(await exo.PreviewMeetingOrganizerChangeAsync(request));
            Assert.Equal(true, runspace.SessionStateProxy.GetVariable("Previewed"));
            Assert.Equal(0, runspace.SessionStateProxy.GetVariable("Transfers"));
            var csv = CalendarTransferRecord.CreateCsvLines(new[]
            {
                new CalendarTransferRecord(DateTimeOffset.UtcNow, request.CurrentOrganizer, request.NewOrganizer,
                    "Subject", "=SUM(1,2)", "Next instance", "Unconfirmed", "Calendar \"outcome\"\nnot verified")
            }).ToArray();
            Assert.Equal(2, csv.Length);
            Assert.Contains("\"'=SUM(1,2)\"", csv[1]);
            Assert.Contains("Unconfirmed", csv[1]);
            Assert.Contains("\"Calendar \"\"outcome\"\"\nnot verified\"", csv[1]);
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData("subject")]
    [InlineData("event")]
    [InlineData("ambiguous")]
    [InlineData("shared")]
    [InlineData("group")]
    [InlineData("new-invalid")]
    [InlineData("same")]
    [InlineData("past")]
    [InlineData("disabled")]
    [InlineData("unexpected")]
    [InlineData("cancelled")]
    public async Task CalendarOrganizerTransferHonorsDocumentedConstraints(string scenario)
    {
        using var runspace = CreateRunspace("""
            $global:Transfers = 0
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                $old = $Identity -like 'old*'
                $type = 'UserMailbox'
                if ($old -and $global:Scenario -eq 'shared') { $type = 'SharedMailbox' }
                if ($old -and $global:Scenario -eq 'group') { $type = 'GroupMailbox' }
                if (!$old -and $global:Scenario -eq 'new-invalid') { $type = 'SharedMailbox' }
                [pscustomobject]@{ PrimarySmtpAddress = $(if ($old -or $global:Scenario -eq 'same') { 'old@example.org' } else { 'new@example.org' }); RecipientTypeDetails = $type }
            }
            function Invoke-ChangeMeetingOrganizer { [CmdletBinding(SupportsShouldProcess)] param($Identity, $NewOrganizer, $EventId, $Subject, [datetime]$TransferSeriesStartDate)
                $global:Transfers++
                if ($Identity -ne 'old@example.org' -or $NewOrganizer -ne 'new@example.org') { throw 'Wrong organizers' }
                if ($PSBoundParameters.ContainsKey('EventId') -eq $PSBoundParameters.ContainsKey('Subject')) { throw 'Exactly one selector required' }
                $global:Selector = $(if ($EventId) { $EventId } else { $Subject })
                $global:UsedEventId = $PSBoundParameters.ContainsKey('EventId')
                $global:TransferDate = $TransferSeriesStartDate
                if ($global:Scenario -eq 'disabled') { throw 'Transfer meeting action is disabled' }
                if ($global:Scenario -eq 'unexpected') { [pscustomobject]@{ Status = 'Unknown' } }
                if ($global:Scenario -eq 'ambiguous') {
                    [pscustomobject]@{ EventId = 'series-1'; Subject = 'Status' }
                    [pscustomobject]@{ EventId = 'series-2'; Subject = 'Status' }
                }
            }
            """);
        runspace.SessionStateProxy.SetVariable("Scenario", scenario);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var request = new MeetingOrganizerRequest("old-alias", "new-alias", scenario == "event", "Status 'review'; $true", DateTime.Today.AddDays(scenario == "past" ? -1 : 2));
            if (scenario == "cancelled")
            {
                exo.OperationCancellationToken = new CancellationToken(true);
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => exo.ChangeMeetingOrganizerAsync(request));
            }
            else if (scenario is "shared" or "group" or "new-invalid" or "same" or "past")
                await Assert.ThrowsAnyAsync<Exception>(() => exo.ChangeMeetingOrganizerAsync(request));
            else
            {
                var validated = await exo.ValidateMeetingOrganizerChangeAsync(request);
                Assert.Equal("old@example.org", validated.CurrentOrganizer);
                Assert.Equal("new@example.org", validated.NewOrganizer);
                Assert.Equal(0, runspace.SessionStateProxy.GetVariable("Transfers"));
                if (scenario is "disabled" or "unexpected")
                    await Assert.ThrowsAnyAsync<Exception>(() => exo.ChangeMeetingOrganizerAsync(request));
                else
                {
                    var matches = await exo.ChangeMeetingOrganizerAsync(request);
                    Assert.Equal(scenario == "ambiguous" ? 2 : 0, matches.Count);
                    Assert.Equal(request.Selector, runspace.SessionStateProxy.GetVariable("Selector"));
                    Assert.Equal(scenario == "event", runspace.SessionStateProxy.GetVariable("UsedEventId"));
                    Assert.Equal(request.TransferSeriesStartDate, runspace.SessionStateProxy.GetVariable("TransferDate"));
                }
            }
            Assert.Equal(scenario is "shared" or "group" or "new-invalid" or "same" or "past" or "cancelled" ? 0 : 1,
                runspace.SessionStateProxy.GetVariable("Transfers"));
        }
        finally { exo.OperationCancellationToken = default; await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData("add")]
    [InlineData("set")]
    [InlineData("none")]
    [InlineData("remove")]
    [InlineData("denied")]
    [InlineData("stale")]
    [InlineData("malformed")]
    [InlineData("delegate")]
    [InlineData("transient-read")]
    [InlineData("transient-write")]
    [InlineData("readback-denied")]
    public async Task CalendarPermissionChangesAreScopedAndVerified(string scenario)
    {
        using var runspace = CreateRunspace("""
            $global:Rights = 'Reviewer'
            $global:LastMutation = ''
            $global:MutationCount = 0
            $global:ReadCount = 0
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                [pscustomobject]@{ PrimarySmtpAddress = 'owner@example.org'; RecipientTypeDetails = 'SharedMailbox' }
            }
            function Get-Recipient { [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                [pscustomobject]@{ PrimarySmtpAddress = 'reader@example.org'; RecipientTypeDetails = 'UserMailbox' }
            }
            function Get-EXOMailboxFolderStatistics { [CmdletBinding()] param($PrimarySmtpAddress, $FolderScope)
                [pscustomobject]@{ FolderType = 'Calendar'; FolderPath = '/Kalender' }
            }
            function Get-EXOMailboxFolderPermission { [CmdletBinding()] param($Identity, $User)
                if ($Identity -ne 'owner@example.org:\Kalender' -or $User -ne 'reader@example.org') { throw 'Wrong permission scope' }
                $global:ReadCount++
                if ($global:Scenario -eq 'transient-read' -and $global:ReadCount -eq 1) { throw 'Object reference not set' }
                if ($global:Scenario -eq 'readback-denied' -and $global:MutationCount -gt 0) { throw 'Read-back access denied' }
                if ($global:Scenario -eq 'denied') { throw 'Access denied' }
                if ($global:Scenario -eq 'malformed') { [pscustomobject]@{ AccessRights = @() }; return }
                if ($null -eq $global:Rights) {
                    $PSCmdlet.WriteError([Management.Automation.ErrorRecord]::new([Exception]::new('No entry'), 'UserNotFoundInPermissionEntryException', [Management.Automation.ErrorCategory]::ObjectNotFound, $User)); return
                }
                [pscustomobject]@{ AccessRights = @($global:Rights); SharingPermissionFlags = @(if ($global:Scenario -eq 'delegate') { 'Delegate' } else { 'None' }) }
            }
            function Add-MailboxFolderPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $User, $AccessRights)
                if ($Identity -ne 'owner@example.org:\Kalender' -or $User -ne 'reader@example.org') { throw 'Wrong mutation scope' }
                $global:MutationCount++
                $global:LastMutation = 'Add'; $global:Rights = $AccessRights
            }
            function Set-MailboxFolderPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $User, $AccessRights)
                if ($Identity -ne 'owner@example.org:\Kalender' -or $User -ne 'reader@example.org') { throw 'Wrong mutation scope' }
                $global:MutationCount++
                $global:LastMutation = 'Set'; $global:Rights = $AccessRights
                if ($global:Scenario -eq 'transient-write') { throw 'Object reference not set' }
            }
            function Remove-MailboxFolderPermission { [CmdletBinding(SupportsShouldProcess)] param($Identity, $User)
                if ($Identity -ne 'owner@example.org:\Kalender' -or $User -ne 'reader@example.org') { throw 'Wrong mutation scope' }
                $global:MutationCount++
                $global:LastMutation = 'Remove'; $global:Rights = $null
            }
            """);
        runspace.SessionStateProxy.SetVariable("Scenario", scenario);
        if (scenario == "add") runspace.SessionStateProxy.SetVariable("Rights", null);
        if (scenario == "delegate") runspace.SessionStateProxy.SetVariable("Rights", "Editor");
        var exo = new ExoPowerShellService(runspace);
        try
        {
            if (scenario is "denied" or "malformed")
                await Assert.ThrowsAnyAsync<Exception>(() => exo.GetCalendarPermissionAsync("owner", "reader"));
            else
            {
                var loaded = await exo.GetCalendarPermissionAsync("owner", "reader");
                if (scenario == "stale")
                {
                    runspace.SessionStateProxy.SetVariable("Rights", "Owner");
                    await Assert.ThrowsAsync<InvalidOperationException>(() => exo.SetCalendarPermissionAsync("owner", "reader", "Editor", loaded));
                }
                else if (scenario == "delegate")
                {
                    await Assert.ThrowsAsync<InvalidOperationException>(() => exo.SetCalendarPermissionAsync("owner", "reader", "Reviewer", loaded));
                    Assert.Contains("Delegate", (await exo.SetCalendarPermissionAsync("owner", "reader", "Editor", loaded)).SharingPermissionFlags);
                }
                else if (scenario == "readback-denied")
                    await Assert.ThrowsAnyAsync<Exception>(() => exo.SetCalendarPermissionAsync("owner", "reader", "Editor", loaded));
                else if (scenario == "remove")
                {
                    Assert.False((await exo.RemoveCalendarPermissionAsync("owner", "reader", loaded)).HasEntry);
                    Assert.Equal("Remove", runspace.SessionStateProxy.GetVariable("LastMutation"));
                }
                else
                {
                    var desired = scenario == "none" ? "None" : "Editor";
                    var changed = await exo.SetCalendarPermissionAsync("owner", "reader", desired, loaded);
                    Assert.True(changed.HasEntry);
                    Assert.Equal(new[] { desired }, changed.AccessRights);
                    Assert.Equal(scenario == "add" ? "Add" : "Set", runspace.SessionStateProxy.GetVariable("LastMutation"));
                }
            }
            if (scenario is "denied" or "malformed" or "stale" or "delegate") Assert.Equal("", runspace.SessionStateProxy.GetVariable("LastMutation"));
            Assert.Equal(scenario is "denied" or "malformed" or "stale" or "delegate" ? 0 : 1, runspace.SessionStateProxy.GetVariable("MutationCount"));
            await Assert.ThrowsAsync<ArgumentException>(() => exo.GetCalendarPermissionAsync("owner", "Default"));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData("localized")]
    [InlineData("denied")]
    [InlineData("missing")]
    [InlineData("ambiguous")]
    [InlineData("malformed")]
    public async Task CalendarFolderDiscoveryIsLocalizedAndFailsClosed(string scenario)
    {
        using var runspace = CreateRunspace("""
            function Get-Mailbox { [CmdletBinding()] param($Identity)
                if ($Identity -ne 'owner-alias') { throw 'Unexpected owner' }
                [pscustomobject]@{ PrimarySmtpAddress = 'owner@example.org'; RecipientTypeDetails = 'SharedMailbox' }
            }
            function Get-EXOMailboxFolderStatistics { [CmdletBinding()] param($PrimarySmtpAddress, $FolderScope)
                if ($PrimarySmtpAddress -ne 'owner@example.org' -or $FolderScope -ne 'Calendar') { throw 'Unscoped folder query' }
                switch ($global:Scenario) {
                    'denied' { throw 'Access denied' }
                    'missing' { return }
                    'ambiguous' { [pscustomobject]@{ FolderType = 'Calendar'; FolderPath = '/Other' } }
                    'malformed' { [pscustomobject]@{ FolderType = 'Calendar' }; return }
                }
                [pscustomobject]@{ FolderType = 'User Created'; FolderPath = '/Secondary' }
                [pscustomobject]@{ FolderType = 'Calendar'; FolderPath = '/Kalender' }
            }
            """);
        runspace.SessionStateProxy.SetVariable("Scenario", scenario);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            if (scenario == "localized") Assert.Equal(@"owner@example.org:\Kalender", await exo.GetCalendarFolderIdentityAsync(" owner-alias "));
            else await Assert.ThrowsAnyAsync<Exception>(() => exo.GetCalendarFolderIdentityAsync("owner-alias"));
            await Assert.ThrowsAsync<ArgumentException>(() => exo.GetCalendarFolderIdentityAsync(" "));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData(ReportKind.GroupMembership, "SharedMailbox", "requires a group")]
    [InlineData(ReportKind.MailboxPermissions, "MailUniversalDistributionGroup", "requires a mailbox")]
    public async Task ReportingSelectionRejectsWrongTargetBeforeReadingReport(ReportKind kind, string recipientType, string expected)
    {
        using var runspace = CreateRunspace("""
            function Get-Recipient { [CmdletBinding()] param($Identity, $RecipientTypeDetails)
                [pscustomobject]@{ Name = 'Target'; Identity = 'canonical-target'; PrimarySmtpAddress = 'target@example.org'; RecipientTypeDetails = $global:TargetType }
            }
            function Get-Mailbox { throw 'Report read must not start' }
            function Get-DistributionGroupMember { throw 'Report read must not start' }
            """);
        runspace.SessionStateProxy.SetVariable("TargetType", recipientType);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var service = new ReportingService(exo);
            var error = await Assert.ThrowsAsync<InvalidOperationException>(() => service.GenerateAsync(kind, "target@example.org"));
            Assert.Contains(expected, error.Message);
            await Assert.ThrowsAsync<ArgumentException>(() => service.GenerateAsync(kind, " "));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => service.GenerateAsync((ReportKind)999, "target@example.org"));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData("found")]
    [InlineData("absent")]
    [InlineData("denied")]
    [InlineData("mismatch")]
    [InlineData("unreadable")]
    [InlineData("ambiguous")]
    public async Task UserSearchMailboxLookupIsScopedAndFailsClosed(string scenario)
    {
        using var runspace = CreateRunspace("""
            function Get-EXOMailbox { [CmdletBinding()] param([Guid]$ExternalDirectoryObjectId)
                if ($ExternalDirectoryObjectId -eq [Guid]::Empty) { throw 'Unscoped query' }
                switch ($global:Scenario) {
                    'absent' { $PSCmdlet.WriteError([Management.Automation.ErrorRecord]::new([Exception]::new('Missing mailbox'), 'ManagementObjectNotFoundException', [Management.Automation.ErrorCategory]::ObjectNotFound, $ExternalDirectoryObjectId)); return }
                    'denied' { throw 'Access denied' }
                    'mismatch' { [pscustomobject]@{ ExternalDirectoryObjectId = [Guid]::NewGuid(); RecipientTypeDetails = 'UserMailbox' }; return }
                    'unreadable' { [pscustomobject]@{ ExternalDirectoryObjectId = $ExternalDirectoryObjectId }; return }
                    'ambiguous' { [pscustomobject]@{ ExternalDirectoryObjectId = $ExternalDirectoryObjectId; RecipientTypeDetails = 'UserMailbox' } }
                }
                [pscustomobject]@{ ExternalDirectoryObjectId = $ExternalDirectoryObjectId; RecipientTypeDetails = 'SharedMailbox' }
            }
            """);
        runspace.SessionStateProxy.SetVariable("Scenario", scenario);
        var exo = new ExoPowerShellService(runspace);
        try
        {
            var objectId = Guid.NewGuid().ToString();
            if (scenario == "found") Assert.Equal("SharedMailbox", await exo.GetUserMailboxTypeAsync(objectId));
            else if (scenario == "absent") Assert.Null(await exo.GetUserMailboxTypeAsync(objectId));
            else await Assert.ThrowsAnyAsync<Exception>(() => exo.GetUserMailboxTypeAsync(objectId));
            await Assert.ThrowsAsync<ArgumentException>(() => exo.GetUserMailboxTypeAsync(""));
        }
        finally { await exo.ShutdownAsync(); }
    }

    [Theory]
    [InlineData("person@example.org", "userPrincipalName eq 'person@example.org' or mail eq 'person@example.org'", null)]
    [InlineData("12345", "employeeId eq 'E12345'", "employeeId eq '12345'")]
    [InlineData("e12345", "employeeId eq 'E12345'", null)]
    [InlineData("O'Neil", "onPremisesSamAccountName eq 'O''Neil'", null)]
    [InlineData("Smith, Jane", "startswith(displayName,'Jane Smith')", null)]
    [InlineData(" Jane   Smith ", "startswith(displayName,'Smith, Jane')", null)]
    public void UserSearchPreservesScoutIdentifiers(string query, string expectedFilter, string? fallback)
    {
        var parsed = Assert.IsType<ParsedSearchQuery>(SearchQueryParser.Parse(query));
        Assert.Contains(expectedFilter, parsed.Filter);
        Assert.Equal(fallback, parsed.FallbackFilter);
    }

    [Fact]
    public void UserSearchDoesNotTurnEmptyInputIntoDirectoryEnumeration()
    {
        Assert.Null(SearchQueryParser.Parse("  "));
        Assert.Throws<ArgumentException>(() => SearchQueryParser.Parse(","));
        Assert.Throws<ArgumentException>(() => SearchQueryParser.Parse("Smith,"));
    }

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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrerequisitesResolvePowerShellHomeWhenCommandLocationHasNoManager(bool brokenFirst)
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $brokenFirst, $documents)
            $savedModulePath = $env:PSModulePath
            try {
                $global:Pwsh = Microsoft.PowerShell.Core\Get-Command pwsh.exe -CommandType Application -ErrorAction Stop
                $global:BrokenFirst = $brokenFirst
                $global:AliasDirectory = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
                function Get-Module { param($Name, [switch]$ListAvailable)
                    if ($Name -like '*\Microsoft.PowerShell.PSResourceGet.psd1') {
                        Microsoft.PowerShell.Core\Get-Module -ListAvailable $Name
                    }
                }
                function Get-Command { param($Name, $CommandType, [switch]$All, $ErrorAction)
                    if ($global:BrokenFirst) { [pscustomobject]@{ Source = (Join-Path $global:AliasDirectory 'pwsh.exe') } }
                    $global:Pwsh
                }
                function Split-Path { param($Path) $global:AliasDirectory }
                function Import-Module { param($Name)
                    if ($Name -like '*\Microsoft.PowerShell.PSResourceGet.psd1') { $global:LoadedManager = $Name; return }
                    throw "Unexpected import: $Name"
                }
                function Get-PSResourceRepository { throw 'TEST: manager resolved' }
                try { & ([scriptblock]::Create($bootstrap)) -DocumentsPath $documents }
                catch { $_.Exception.Message }
                $global:LoadedManager
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var documents = Path.Combine(Path.GetTempPath(), "EXOKit-docs-" + Guid.NewGuid().ToString("N"));
        try
        {
            powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("brokenFirst", brokenFirst).AddParameter("documents", documents);
            var results = powershell.Invoke();
            Assert.Empty(powershell.Streams.Error);
            Assert.Equal(2, results.Count);
            Assert.Equal("TEST: manager resolved", results[0].ToString());
            Assert.EndsWith("Microsoft.PowerShell.PSResourceGet.psd1", results[1].ToString());
        }
        finally { if (Directory.Exists(documents)) Directory.Delete(documents, true); }
    }

    [Fact]
    public void PrerequisitesDistinguishMissingManagerFromEmbeddedEngineVersion()
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $documents, $localData)
            $savedModulePath = $env:PSModulePath
            try {
                function Get-Module { param($Name, [switch]$ListAvailable) }
                function Get-Command { param($Name, $CommandType, [switch]$All, $ErrorAction) }
                function Test-Path { param($LiteralPath, $PathType) $false }
                function Invoke-WebRequest { throw 'offline' }
                try { & ([scriptblock]::Create($bootstrap)) -DocumentsPath $documents -LocalDataPath $localData }
                catch { $_.Exception.Message }
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var documents = Path.Combine(Path.GetTempPath(), "EXOKit-docs-" + Guid.NewGuid().ToString("N"));
        powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("documents", documents).AddParameter("localData", documents + "-local");
        List<PSObject> results;
        try { results = powershell.Invoke().ToList(); }
        finally
        {
            if (Directory.Exists(documents)) Directory.Delete(documents, true);
            if (Directory.Exists(documents + "-local")) Directory.Delete(documents + "-local", true);
        }
        Assert.Empty(powershell.Streams.Error);
        var message = Assert.Single(results).ToString();
        Assert.Contains("cannot load Microsoft.PowerShell.PSResourceGet 1.2.0 or later (embedded PowerShell 7.6", message);
        Assert.Contains("Private PSResourceGet download: offline", message);
        Assert.Contains("-Scope CurrentUser", message);
        Assert.DoesNotContain("Install PowerShell 7.6", message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrerequisitesSkipUnloadablePowerShellCopyThenVerifyPrivateDownload(bool tamperedDownload)
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $documents, $localData)
            $savedModulePath = $env:PSModulePath
            try {
                $storeCopy = 'C:\Program Files\WindowsApps\Microsoft.PowerShell_7.6.6.0_x64__8wekyb3d8bbwe\Modules\Microsoft.PowerShell.PSResourceGet\Microsoft.PowerShell.PSResourceGet.psd1'
                function Get-Module { param($Name, [switch]$ListAvailable) [pscustomobject]@{ Version = [version]'1.2.0'; Path = $storeCopy } }
                function Get-Command { param($Name, $CommandType, [switch]$All, $ErrorAction) }
                function Test-Path { param($LiteralPath, $PathType) $false }
                function Import-Module { param($Name) throw 'Could not load file or assembly. Access is denied.' }
                function Invoke-WebRequest { param($Uri, $OutFile) [IO.File]::WriteAllText($OutFile, 'not the published package') }
                try { & ([scriptblock]::Create($bootstrap)) -DocumentsPath $documents -LocalDataPath $localData }
                catch { $_.Exception.Message }
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var root = Path.Combine(Path.GetTempPath(), "EXOKit-store-" + Guid.NewGuid().ToString("N"));
        powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("documents", Path.Combine(root, "Documents")).AddParameter("localData", Path.Combine(root, "Local"));
        try
        {
            var message = Assert.Single(powershell.Invoke()).ToString();
            Assert.Empty(powershell.Streams.Error);
            Assert.Contains("WindowsApps", message);
            Assert.Contains("Access is denied", message);
            Assert.Contains("failed integrity verification", message);
            Assert.False(Directory.Exists(Path.Combine(root, "Local", "EXOKit", "PSResourceGet", "1.2.0")));
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public void PrerequisitesInstallVerifiedPrivatePsResourceGetWhenNoPowerShellInstallExists()
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $documents, $localData, $empty)
            $savedModulePath = $env:PSModulePath; $savedProgramFiles = $env:ProgramFiles; $savedProgramW6432 = $env:ProgramW6432
            try {
                $env:ProgramFiles = $empty; $env:ProgramW6432 = $empty
                function Get-Command { param($Name, $CommandType, [switch]$All, $ErrorAction) }
                function Get-Module { param($Name, [switch]$ListAvailable)
                    if ($Name -like '*.psd1') { Microsoft.PowerShell.Core\Get-Module -ListAvailable $Name }
                }
                function Get-PSResourceRepository { throw 'TEST: manager loaded' }
                try { & ([scriptblock]::Create($bootstrap)) -DocumentsPath $documents -LocalDataPath $localData }
                catch { $_.Exception.Message }
                (Microsoft.PowerShell.Core\Get-Module Microsoft.PowerShell.PSResourceGet).Path
            } finally { $env:PSModulePath = $savedModulePath; $env:ProgramFiles = $savedProgramFiles; $env:ProgramW6432 = $savedProgramW6432 }
            """);
        var root = Path.Combine(Path.GetTempPath(), "EXOKit-private-" + Guid.NewGuid().ToString("N"));
        powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("documents", Path.Combine(root, "Documents"))
            .AddParameter("localData", Path.Combine(root, "Local")).AddParameter("empty", Path.Combine(root, "NoPowerShell"));
        try
        {
            var results = powershell.Invoke();
            Assert.Empty(powershell.Streams.Error);
            Assert.Equal("TEST: manager loaded", results[0].ToString());
            var expected = Path.Combine(root, "Local", "EXOKit", "PSResourceGet", "1.2.0");
            Assert.Equal(expected, Path.GetDirectoryName(results[1].ToString()), ignoreCase: true);
            var folders = Directory.GetDirectories(Path.Combine(root, "Local", "EXOKit", "PSResourceGet")).Select(Path.GetFileName);
            Assert.Equal(new[] { "1.2.0" }, folders);
        }
        finally
        {
            runspace.Close();
            if (Directory.Exists(root)) try { Directory.Delete(root, true); } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrerequisitesFallBackWhenDocumentsModulesFolderIsBlocked(bool fallbackAlsoBlocked)
    {
        using var runspace = RunspaceFactory.CreateRunspace(new LoggerPSHost(), LoggerPSHost.CreateInitialSessionState());
        runspace.Open();
        using var powershell = PowerShell.Create();
        powershell.Runspace = runspace;
        powershell.AddScript("""
            param($bootstrap, $documents, $localData)
            $savedModulePath = $env:PSModulePath
            try {
                function Get-Module { param($Name, [switch]$ListAvailable) }
                function Get-Command { param($Name, $CommandType, [switch]$All, $ErrorAction) }
                function Test-Path { param($LiteralPath, $PathType) $false }
                function Invoke-WebRequest { throw 'offline' }
                $message = try { & ([scriptblock]::Create($bootstrap)) -DocumentsPath $documents -LocalDataPath $localData } catch { $_.Exception.Message }
                [pscustomobject]@{ Message = $message; ModulePath = $env:PSModulePath }
            } finally { $env:PSModulePath = $savedModulePath }
            """);
        var root = Path.Combine(Path.GetTempPath(), "EXOKit-blocked-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var blockingFile = Path.Combine(root, "blocked");
        File.WriteAllText(blockingFile, "a file where a directory is required");
        var documents = Path.Combine(blockingFile, "Documents");
        var localData = fallbackAlsoBlocked ? Path.Combine(blockingFile, "Local") : Path.Combine(root, "Local");
        powershell.AddParameter("bootstrap", PowerShellPrerequisites.Script).AddParameter("documents", documents).AddParameter("localData", localData);
        try
        {
            var result = Assert.Single(powershell.Invoke());
            Assert.Empty(powershell.Streams.Error);
            var message = result.Properties["Message"].Value?.ToString() ?? string.Empty;
            if (fallbackAlsoBlocked)
            {
                Assert.Contains("cannot write PowerShell modules", message);
                Assert.Contains("Controlled Folder Access", message);
                Assert.DoesNotContain(Path.Combine(localData, "EXOKit", "Modules") + ";", result.Properties["ModulePath"].Value?.ToString());
            }
            else
            {
                var fallback = Path.Combine(localData, "EXOKit", "Modules");
                Assert.True(Directory.Exists(fallback));
                Assert.Empty(Directory.GetFileSystemEntries(fallback));
                Assert.Contains(fallback, result.Properties["ModulePath"].Value?.ToString());
                Assert.Contains(powershell.Streams.Warning, warning => warning.Message.Contains(fallback));
                Assert.Contains("cannot load Microsoft.PowerShell.PSResourceGet", message);
            }
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ShellBrushesAreDefinedForEveryApplicationTheme()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "App.xaml"))) directory = directory.Parent;
        Assert.NotNull(directory);
        var document = System.Xml.Linq.XDocument.Load(Path.Combine(directory!.FullName, "App.xaml"));
        System.Xml.Linq.XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
        var themes = document.Descendants().Where(element => element.Name.LocalName == "ResourceDictionary.ThemeDictionaries")
            .SelectMany(element => element.Elements()).ToDictionary(element => (string)element.Attribute(x + "Key")!,
                element => element.Elements().Select(brush => (string?)brush.Attribute(x + "Key")).OfType<string>().ToHashSet());
        // Windows resolves a missing Light theme through Default; the dark shell must never fail on a light-mode PC.
        Assert.True(themes.ContainsKey("Default") || themes.ContainsKey("Light"), "No theme dictionary serves light-mode Windows.");
        var shellBrushes = themes.Values.SelectMany(keys => keys).Where(key => key.StartsWith("Shell", StringComparison.Ordinal)).ToHashSet();
        Assert.NotEmpty(shellBrushes);
        foreach (var theme in new[] { "Default", "Dark", "HighContrast" })
            Assert.True(themes.TryGetValue(theme, out var keys) && shellBrushes.IsSubsetOf(keys), $"{theme} theme is missing shell brushes.");
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
            var report = await new ReportingService(exo).GenerateAsync(ReportKind.MailboxPermissions, "mailbox@example.org");
            Assert.Equal(ReportTargetKind.Mailbox, report.Target.Kind);
            var rows = report.Rows;
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
            var report = await new ReportingService(exo).GenerateAsync(ReportKind.GroupMembership, "group");
            Assert.Equal(ReportTargetKind.Group, report.Target.Kind);
            var rows = report.Rows;
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

    private sealed class GraphHandler(params string[] responses) : HttpMessageHandler
    {
        public string? Request { get; private set; }
        public List<string> Requests { get; } = new();
        public List<bool> EventualHeaders { get; } = new();
        public System.Net.HttpStatusCode StatusCode { get; init; } = System.Net.HttpStatusCode.OK;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Request = Uri.UnescapeDataString(request.RequestUri!.ToString());
            Requests.Add(Request);
            EventualHeaders.Add(request.Headers.TryGetValues("ConsistencyLevel", out var values) && values.Contains("eventual"));
            return Task.FromResult(new HttpResponseMessage(StatusCode)
            {
                Content = new StringContent(responses[Math.Min(Requests.Count - 1, responses.Length - 1)], System.Text.Encoding.UTF8, "application/json")
            });
        }
    }

    [Fact]
    public async Task UserSearchReadsAllPagesAndPreservesDeletedAndUnknownAccountState()
    {
        using var handler = new GraphHandler(
            """{"value":[{"id":"one","displayName":"Zoe","accountEnabled":true}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/users?$skiptoken=next"}""",
            """{"value":[{"id":"two","displayName":"Amy","employeeId":"E123","onPremisesSamAccountName":"amy","businessPhones":["111","222"]}]}""",
            """{"value":[{"id":"deleted","displayName":"Deleted user"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$skiptoken=next"}""",
            """{"value":[{"id":"deleted-two","displayName":"Former user"}]}""");
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var service = new GraphService(new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", ""), client);
        var results = await service.SearchUsersAsync("Amy", true);
        Assert.Equal(4, results.Count);
        Assert.Equal("Amy", results[0].DisplayName);
        Assert.Equal("Unknown", results[0].AccountStatus);
        Assert.Equal("111; 222", results[0].BusinessPhone);
        Assert.Equal(2, results.Count(user => user.IsDeleted));
        Assert.All(handler.EventualHeaders, Assert.True);
        Assert.Contains("$count=true", handler.Requests[0]);
        Assert.Contains("$skiptoken=next", handler.Requests[1]);
        Assert.Contains("deletedItems/graph.user", handler.Requests[2]);
    }

    [Fact]
    public async Task UserSearchRetriesNumericEmployeeIdSeparatelyForActiveAndDeletedUsers()
    {
        using var handler = new GraphHandler("""{"value":[]}""", """{"value":[{"id":"active"}]}""", """{"value":[]}""", """{"value":[{"id":"deleted"}]}""");
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var service = new GraphService(new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", ""), client);
        Assert.Equal(2, (await service.SearchUsersAsync("123", true)).Count);
        Assert.Contains("employeeId eq 'E123'", handler.Requests[0]);
        Assert.Contains("employeeId eq '123'", handler.Requests[1]);
        Assert.Contains("employeeId eq 'E123'", handler.Requests[2]);
        Assert.Contains("employeeId eq '123'", handler.Requests[3]);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"value\":[{}]}")]
    public async Task UserSearchRejectsUnreadableResponses(string response)
    {
        using var handler = new GraphHandler(response);
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var service = new GraphService(new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", ""), client);
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.SearchUsersAsync("Amy"));
    }

    [Fact]
    public async Task UserSearchPropagatesPermissionErrorsAndCancellation()
    {
        using var handler = new GraphHandler("""{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}""")
            { StatusCode = System.Net.HttpStatusCode.Forbidden };
        using var http = new HttpClient(handler);
        using var client = new Microsoft.Graph.GraphServiceClient(http, new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider());
        var auth = new AuthService(Array.Empty<string>(), () => IntPtr.Zero, "", "");
        var service = new GraphService(auth, client);
        await Assert.ThrowsAsync<Microsoft.Graph.Models.ODataErrors.ODataError>(() => service.SearchUsersAsync("Amy", true));
        auth.OperationCancellationToken = new CancellationToken(true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => service.SearchUsersAsync("Amy"));
        Assert.Single(handler.Requests);
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