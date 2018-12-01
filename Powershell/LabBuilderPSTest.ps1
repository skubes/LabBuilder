using namespace LabBuilder
$assName = "C:\Users\steve\Source\Workspaces\LabUtils\Dev\LabBuilder\LabBuilder.exe\Debug\LabBuilder.exe"
$env:BUILD_SOURCESDIRECTORY | Out-File -FilePath .\LabuildBuilderPowerShellTest.txt
return
[Reflection.Assembly]::LoadFile($assName)  | Out-Null

$logFile = [Gen]::InitializeFileLogging()
Write-Host "Log file located: [$logFile]"

Write-Host "Calling Gen::Mail"
$GenMail = [Gen]::Mail

$mailArgs = New-Object GenMailArgs
$mailArgs.InitialItems = 1
$mailArgs.SelectedOUPath = "LDAP://ou=ev122,dc=hillvalley,dc=com"
$mailArgs.Threads = 32
$mailArgs.PercentChanceOfExtraRecipients = 20
$mailArgs.PercentChanceOfAttachments = 20
$mailArgs.MaxAttachments = 5
$mailArgs.MaxAdditionalRecipients = 10

$GenMail.Invoke($mailArgs)
Write-Host "Complete"


