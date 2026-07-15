$ErrorActionPreference = 'Stop'

# Fill these in for the release you're publishing, then `choco pack` from
# the packaging/chocolatey folder. Get the checksum from that release's
# SHA256SUMS file (see README: Verifying a release).
$url      = 'https://github.com/Yaksharo/appraisal-filler/releases/download/vX.Y/AdviseeDocFiller-Setup-X.Y.exe'
$checksum = 'PASTE_SHA256_FROM_RELEASE_SHA256SUMS_HERE'

$packageArgs = @{
  packageName    = 'adviseedocfiller'
  fileType       = 'exe'
  url            = $url
  checksum       = $checksum
  checksumType   = 'sha256'
  silentArgs     = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
  validExitCodes = @(0)
}

Install-ChocolateyPackage @packageArgs
