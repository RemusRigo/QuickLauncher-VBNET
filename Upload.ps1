param([string]$msg)

if (-not $msg)
{
   $msg = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

if (!(Test-Path ".git"))
{
   git init
   git branch -M main
   git remote add origin "https://github.com/RemusRigo/QuickLauncher-VBNET.git"
}

git add .
git commit -m "$msg"
git push -u origin main
