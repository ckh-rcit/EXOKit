namespace EXOKit.Services;

internal static class PowerShellPrerequisites
{
    internal const string Script = """
        param([string]$ModuleRoot)
        $ErrorActionPreference = 'Stop'
        Write-Host "EXOKit embedded PowerShell: $($PSVersionTable.PSVersion); PSHOME: $PSHOME"
        if ($PSVersionTable.PSVersion -lt [version]'7.6.0') { throw "Update EXOKit: EXO 3.10.1 requires embedded PowerShell 7.6 or later; this application is using $($PSVersionTable.PSVersion). Updating standalone PowerShell does not update the application's embedded engine." }
        $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
        if ([string]::IsNullOrWhiteSpace($documents)) { throw 'The CurrentUser Documents directory is unavailable. Cannot prepare PowerShell modules.' }
        $userModules = Join-Path $documents 'PowerShell\Modules'
        if ($ModuleRoot) { $userModules = $ModuleRoot }
        if (($env:PSModulePath -split [IO.Path]::PathSeparator) -notcontains $userModules) {
            $env:PSModulePath = $userModules + [IO.Path]::PathSeparator + $env:PSModulePath
        }
        $manager = Get-Module -ListAvailable Microsoft.PowerShell.PSResourceGet |
            Where-Object Version -GE ([version]'1.2.0') | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $manager) {
            $pwshPaths = @(
                (Get-Command pwsh.exe -CommandType Application -All -ErrorAction SilentlyContinue).Source
                foreach ($programFiles in @($env:ProgramW6432, $env:ProgramFiles)) {
                    if ($programFiles) {
                        $installedPwsh = Join-Path $programFiles 'PowerShell\7\pwsh.exe'
                        if (Test-Path -LiteralPath $installedPwsh -PathType Leaf) { $installedPwsh }
                    }
                }
            ) | Where-Object { $_ } | Select-Object -Unique
            $discoveryErrors = [Collections.Generic.List[string]]::new()
            foreach ($pwshPath in $pwshPaths) {
                $process = $null
                try {
                    $manifest = Join-Path (Split-Path $pwshPath) 'Modules\Microsoft.PowerShell.PSResourceGet\Microsoft.PowerShell.PSResourceGet.psd1'
                    if (Test-Path -LiteralPath $manifest -PathType Leaf) {
                        $manager = Get-Module -ListAvailable $manifest | Where-Object Version -GE ([version]'1.2.0') | Select-Object -First 1
                    }
                    if (-not $manager) {
                        $process = [Diagnostics.Process]::new()
                        $process.StartInfo.FileName = $pwshPath
                        $process.StartInfo.UseShellExecute = $false
                        $process.StartInfo.CreateNoWindow = $true
                        $process.StartInfo.RedirectStandardOutput = $true
                        $process.StartInfo.RedirectStandardError = $true
                        foreach ($argument in @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', '$PSHOME')) {
                            $process.StartInfo.ArgumentList.Add($argument)
                        }
                        $process.Start() | Out-Null
                        $stdout = $process.StandardOutput.ReadToEndAsync()
                        $stderr = $process.StandardError.ReadToEndAsync()
                        if (-not $process.WaitForExit(10000)) {
                            $process.Kill($true)
                            throw 'PowerShell home discovery timed out after 10 seconds.'
                        }
                        $powerShellHome = $stdout.GetAwaiter().GetResult().Trim()
                        $processError = $stderr.GetAwaiter().GetResult().Trim()
                        if ($process.ExitCode -ne 0 -or [string]::IsNullOrWhiteSpace($powerShellHome)) {
                            throw "PowerShell home discovery failed (exit $($process.ExitCode)): $processError"
                        }
                        $manifest = Join-Path $powerShellHome 'Modules\Microsoft.PowerShell.PSResourceGet\Microsoft.PowerShell.PSResourceGet.psd1'
                        if (Test-Path -LiteralPath $manifest -PathType Leaf) {
                            $manager = Get-Module -ListAvailable $manifest | Where-Object Version -GE ([version]'1.2.0') | Select-Object -First 1
                        }
                    }
                    if ($manager) { break }
                    $discoveryErrors.Add("$pwshPath : PSResourceGet 1.2.0 or later was not found.")
                } catch {
                    $discoveryErrors.Add("$pwshPath : $($_.Exception.Message)")
                } finally {
                    if ($process) { $process.Dispose() }
                }
            }
            if (-not $manager) { throw "EXOKit is running embedded PowerShell $($PSVersionTable.PSVersion), but cannot locate Microsoft.PowerShell.PSResourceGet 1.2.0 or later. In standalone PowerShell 7, run 'Install-Module Microsoft.PowerShell.PSResourceGet -MinimumVersion 1.2.0 -Scope CurrentUser -Repository PSGallery', then restart EXOKit. Discovery details: $($discoveryErrors -join ' ')" }
        }
        Write-Host "PSResourceGet $($manager.Version): $($manager.Path)"
        Import-Module $manager.Path
        $repository = Get-PSResourceRepository PSGallery -ErrorAction SilentlyContinue
        if (-not $repository) { Register-PSResourceRepository -PSGallery; $repository = Get-PSResourceRepository PSGallery }
        if ($repository.Uri.ToString().TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') { throw 'PSGallery must point to https://www.powershellgallery.com/api/v2.' }
        $requirements = [ordered]@{ PackageManagement = [version]'1.4.8.1'; PowerShellGet = [version]'2.2.5'; ExchangeOnlineManagement = [version]'3.10.1' }
        foreach ($name in $requirements.Keys) {
            $installed = Get-Module -ListAvailable $name | Where-Object { -not $_.PrivateData.PSData.Prerelease } |
                Sort-Object Version -Descending | Select-Object -First 1
            $version = $requirements[$name]
            if ($installed -and $installed.Version -gt $version) { $version = $installed.Version }
            try {
                $latest = Find-PSResource $name -Repository PSGallery
                if ($latest -and [version]$latest.Version -gt $version) { $version = [version]$latest.Version }
            } catch {
                if (-not $installed -or $installed.Version -lt $requirements[$name]) { throw }
                Write-Warning "Cannot check updates for $name. Using installed $($installed.Version)."
            }
            $loaded = Get-Module $name
            if ($loaded -and @($loaded | Where-Object Version -NE $version).Count) { throw "$name has an older version loaded. Restart the application before updating modules." }
            if (-not $installed -or $installed.Version -ne $version) {
                Write-Host "Installing $name $version for CurrentUser..."
                [IO.Directory]::CreateDirectory($userModules) | Out-Null
                Save-PSResource $name -Version $version.ToString() -Repository PSGallery -Path $userModules -IncludeXml -TrustRepository -Quiet
                $modulePath = Join-Path $userModules "$name\$version\$name.psd1"
            } else {
                $modulePath = $installed.Path
            }
            if (-not (Test-Path -LiteralPath $modulePath)) { throw "$name $version was not saved at $modulePath. Module preparation is incomplete." }
            Import-Module $modulePath -RequiredVersion $version -Global
            Write-Host "Ready: $name $version"
        }
        foreach ($name in @('Connect-ExchangeOnline','Disconnect-ExchangeOnline','Get-ConnectionInformation','Get-EXOMailbox','Update-ModuleManifest')) {
            if (-not (Get-Command $name -ErrorAction SilentlyContinue)) { throw "Required cmdlet missing: $name. Repair the module and restart the application." }
        }
        """;
}