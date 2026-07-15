$ErrorActionPreference = 'Stop'

$uninstallPath = Join-Path ${env:ProgramFiles} 'Yaksharo Solutions\Advisee Document Filler\unins000.exe'

if (Test-Path $uninstallPath) {
  Uninstall-ChocolateyPackage -PackageName 'adviseedocfiller' -FileType 'exe' `
    -SilentArgs '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-' `
    -File $uninstallPath -ValidExitCodes @(0)
}
