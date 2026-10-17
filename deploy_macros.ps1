param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Dev", "QA", "Prod")]
    [string]$Environment
)

# Define Base Paths (Adjust C:\Environments if your server target path differs)
$repoBaseDir   = "C:\Users\HP\OneDrive\Projects\VBA"
$targetBaseDir = "D:\ExcelAutomations\Macros\Environments\$Environment"

# Ensure target environment folder exists
if (!(Test-Path $targetBaseDir)) { 
    New-Item -ItemType Directory -Path $targetBaseDir -Force 
}

# Map Environment to corresponding Git branch
switch ($Environment) {
    "Dev"  { $branch = "dev" }
    "QA"   { $branch = "qa" }
    "Prod" { $branch = "main" }
}

Write-Host "Fetching latest code for branch [$branch]..." -ForegroundColor Cyan

# 1. Pull latest changes from GitHub
Set-Location $repoBaseDir
git fetch origin
git checkout $branch
git pull origin $branch

Write-Host "Deploying files to environment [$Environment] at [$targetBaseDir]..." -ForegroundColor Green

# 2. Recursively copy all folders (excel, vba, etc.) to the target environment directory
# Excludes the hidden .git metadata folder and gitignore from being copied over
Get-ChildItem -Path $repoBaseDir -Exclude ".git", ".gitignore", "dev_env", ".venv" | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $targetBaseDir -Recurse -Force
}

# 3. Copy environment-specific config file as active config.json if it exists
if (Test-Path "$repoBaseDir\config.$Environment.json") {
    Copy-Item -Path "$repoBaseDir\config.$Environment.json" -Destination "$targetBaseDir\config.json" -Force
}

Write-Host "Deployment to [$Environment] completed successfully!" -ForegroundColor Green