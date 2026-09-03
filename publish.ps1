# ==============================================================================
# Automate Build, Git Push, and NPM Package Publication
# Usage: .\publish.ps1
# ==============================================================================
$ErrorActionPreference = "Stop"

Clear-Host
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   spsignature Automated Publisher Script" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# Read current version from package.json
$packageJson = Get-Content -Path "package.json" -Raw | ConvertFrom-Json
$currentVersion = $packageJson.version

# Parse version numbers to show exact previews
$versionParts = $currentVersion -split '\.'
$major = [int]$versionParts[0]
$minor = [int]$versionParts[1]
$patch = [int]$versionParts[2]

$nextPatch = "$major.$minor.$($patch + 1)"
$nextMinor = "$major.$($minor + 1).0"
$nextMajor = "$($major + 1).0.0"

# Helper function
function Invoke-ExternalCommand {
    param (
        [scriptblock]$Command,
        [string]$ErrorMessage
    )
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw $ErrorMessage
    }
}

# Check NPM auth
Write-Host ""
Write-Host "==> Checking NPM Authentication Status..." -ForegroundColor Cyan
$npmUser = & npm whoami 2>$null
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($npmUser)) {
    Write-Host ""
    Write-Host "WARNING: You are NOT logged into NPM!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Your authentication token is missing, expired, or invalid." -ForegroundColor Yellow
    Write-Host "You must run 'npm login' before publishing." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Run this command in your terminal:" -ForegroundColor Cyan
    Write-Host "  npm login" -ForegroundColor White
    Write-Host ""
    Write-Host "After logging in, run this script again." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
else {
    Write-Host "Authenticated as: $npmUser" -ForegroundColor Green
}

# Version selection
Write-Host ""
Write-Host "Select version bump type:" -ForegroundColor Yellow
Write-Host "  [1] patch (e.g. $currentVersion -> $nextPatch) - Default"
Write-Host "  [2] minor (e.g. $currentVersion -> $nextMinor)"
Write-Host "  [3] major (e.g. $currentVersion -> $nextMajor)"
Write-Host "  [0] none  (Keep $currentVersion - Use if retrying after a failure)"
$bumpInput = Read-Host "Choice [1]"

$BumpType = "patch"
if ($bumpInput -eq "2") { $BumpType = "minor" }
elseif ($bumpInput -eq "3") { $BumpType = "major" }
elseif ($bumpInput -eq "0") { $BumpType = "none" }

# Target action selection
$doGit = $true
$doNpm = $true

if ($BumpType -eq "none") {
    Write-Host ""
    Write-Host "Select Target Action:" -ForegroundColor Yellow
    Write-Host "  [1] Both (Git Push and NPM Publish) - Default"
    Write-Host "  [2] Git Only (Commit and Push to GitHub only)"
    Write-Host "  [3] NPM Only (Build and Publish to NPM only)"
    $actionInput = Read-Host "Choice [1]"

    if ($actionInput -eq "2") {
        $doGit = $true
        $doNpm = $false
    }
    elseif ($actionInput -eq "3") {
        $doGit = $false
        $doNpm = $true
    }
}

# Build
Write-Host "`n==> 1. Running TypeScript Compilation..." -ForegroundColor Cyan
Invoke-ExternalCommand -Command { npm run build } -ErrorMessage "TypeScript compilation failed."

# Version bump
if ($BumpType -eq "none") {
    $newVersion = $currentVersion
    Write-Host "`n==> 2. Keeping existing version ($newVersion)..." -ForegroundColor Yellow
}
else {
    Write-Host "`n==> 2. Bumping Package Version ($BumpType)..." -ForegroundColor Cyan
    $newVersionRaw = & npm version $BumpType --no-git-tag-version
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to bump package version."
    }
    $newVersion = $newVersionRaw.Trim()
    Write-Host "New Version: $newVersion" -ForegroundColor Green
}

# Git sync
if ($doGit) {
    Write-Host "`n==> 3. Checking Git Working Tree..." -ForegroundColor Cyan
    Invoke-ExternalCommand -Command { git add . } -ErrorMessage "Failed to stage git files."

    $gitStatus = git status --porcelain
    if ($gitStatus) {
        Write-Host ""
        $CommitMessage = ""
        while ([string]::IsNullOrWhiteSpace($CommitMessage)) {
            $CommitMessage = Read-Host "Enter Git Commit Message (Required)"
            if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
                Write-Host "Error: Commit message cannot be empty!" -ForegroundColor Red
            }
        }

        $fullCommitMessage = "$CommitMessage ($newVersion)"
        Invoke-ExternalCommand -Command { git commit -m $fullCommitMessage } -ErrorMessage "Git commit failed."
        Write-Host "Committed changes: $fullCommitMessage" -ForegroundColor Green
    }
    else {
        Write-Host "No uncommitted file changes detected. Skipping commit step." -ForegroundColor Yellow
    }

    Write-Host "Pushing to remote repository (main)..." -ForegroundColor Gray
    try {
        Invoke-ExternalCommand -Command { git push origin main } -ErrorMessage "Git push failed."
        Write-Host "GitHub synchronization complete!" -ForegroundColor Green
    }
    catch {
        Write-Host "`n[ERROR] Git push failed." -ForegroundColor Red
        Write-Host "Fix the issue and rerun with option [0] to retry." -ForegroundColor Yellow
        exit 1
    }
}
else {
    Write-Host "`n==> 3. Git synchronization skipped by user." -ForegroundColor Yellow
}

# NPM publish
if ($doNpm) {
    Write-Host "`n==> 4. NPM Registry Publication..." -ForegroundColor Cyan
    $confirmPublish = Read-Host "Ready to publish $newVersion to NPM? [Y/n]"
    if ($confirmPublish -eq "" -or $confirmPublish.ToUpper() -eq "Y") {
        try {
            Invoke-ExternalCommand -Command { npm publish --access public } -ErrorMessage "NPM publish failed."
            Write-Host "Successfully published $newVersion to NPM!" -ForegroundColor Green
        }
        catch {
            Write-Host "`n[ERROR] NPM publish failed." -ForegroundColor Red
            Write-Host "Run 'npm login' and rerun with option [0]." -ForegroundColor Yellow
            exit 1
        }
    }
    else {
        Write-Host "NPM Publish skipped by user." -ForegroundColor Yellow
    }
}
else {
    Write-Host "`n==> 4. NPM Publication skipped by user." -ForegroundColor Yellow
}

Write-Host "`nAll operations completed successfully!" -ForegroundColor Green