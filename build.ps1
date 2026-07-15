$ErrorActionPreference = 'Stop'
$Project = Split-Path -Parent $MyInvocation.MyCommand.Path
$Output = Join-Path $Project 'outputs\EmailAutomationDeployment'
$Work = Join-Path $Project 'work\pyinstaller'
New-Item -ItemType Directory -Path $Output -Force | Out-Null
New-Item -ItemType Directory -Path $Work -Force | Out-Null
Remove-Item (Join-Path $Output '*.exe') -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $Output '.env') -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $Output '.env.template') -Force -ErrorAction SilentlyContinue
$Names = @('Email Automation')
foreach ($Name in $Names) {
    & 'C:\Python314\python.exe' -m PyInstaller --noconfirm --clean --onefile --windowed --name $Name --distpath $Output --workpath (Join-Path $Work $Name) --specpath $Work (Join-Path $Project 'email_automation.py')
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed for $Name" }
}
Copy-Item (Join-Path $Project 'README.md') $Output -Force
foreach ($Folder in @('config','backup')) { New-Item -ItemType Directory -Path (Join-Path $Output $Folder) -Force | Out-Null }
& 'C:\Python314\python.exe' -c "import sys; sys.path.insert(0, r'$Project'); from email_automation import sample_workbook; sample_workbook(r'$Output\sample_mail_list.xlsx')"
if ($LASTEXITCODE -ne 0) { throw 'Sample workbook creation failed' }
