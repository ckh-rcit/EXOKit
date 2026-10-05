namespace EXOKit.Services;

internal static class PowerShellPrerequisites
{
    internal const string Script = """
        $ErrorActionPreference = 'Stop'
        if ($PSVersionTable.PSVersion -lt [version]'7.6.0') { throw 'Update the application: EXO 3.10.1 requires embedded PowerShell 7.6 or later.' }
        $manager = Get-Module -ListAvailable Microsoft.PowerShell.PSResourceGet |
            Where-Object Version -GE ([version]'1.2.0') | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $manager) {
            $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwsh) {
                $manifest = Join-Path (Split-Path $pwsh.Source) 'Modules\Microsoft.PowerShell.PSResourceGet\Microsoft.PowerShell.PSResourceGet.psd1'
                $manager = Get-Module -ListAvailable $manifest | Where-Object Version -GE ([version]'1.2.0') | Select-Object -First 1
            }
        }
        if (-not $manager) { throw 'Install PowerShell 7.6 or later, which includes PSResourceGet, then restart this application.' }
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
                Install-PSResource $name -Version $version.ToString() -Repository PSGallery -Scope CurrentUser -TrustRepository -Quiet
            }
            Import-Module $name -RequiredVersion $version -Global
            Write-Host "Ready: $name $version"
        }
        foreach ($name in @('Connect-ExchangeOnline','Disconnect-ExchangeOnline','Get-ConnectionInformation','Get-EXOMailbox','Update-ModuleManifest')) {
            if (-not (Get-Command $name -ErrorAction SilentlyContinue)) { throw "Required cmdlet missing: $name. Repair the module and restart the application." }
        }
        """;
}