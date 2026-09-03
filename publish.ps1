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

# ------------------------------------------------------------------------------
# 1. Version Selection (Includes "none" for retries)
# ------------------------------------------------------------------------------
Write-Host ""
Write-Host "Current Package Version: $currentVersion" -ForegroundColor Green
Write-Host "Select version bump type:" -ForegroundColor Yellow
Write-Host "  [1] patch (e.g. $currentVersion -> patch) - Default"
Write-Host "  [2] minor (e.g. $currentVersion -> minor)"
Write-Host "  [3] major (e.g. $currentVersion -> major)"
Write-Host "  [0] none  (Keep $currentVersion - Use if retrying after a failure)"
$bumpInput = Read-Host "Choice [1]"

$BumpType = "patch"
if ($bumpInput -eq "2") { $BumpType = "minor" }
elseif ($bumpInput -eq "3") { $BumpType = "major" }
elseif ($bumpInput -eq "0") { $BumpType = "none" }

# ------------------------------------------------------------------------------
# Helper: Run external process and enforce exit-code safety
# ------------------------------------------------------------------------------
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

# ------------------------------------------------------------------------------
# 2. Build Pipeline
# ------------------------------------------------------------------------------
Write-Host "`n==> 1. Running TypeScript Compilation..." -ForegroundColor Cyan
Invoke-ExternalCommand -Command { npm run build } -ErrorMessage "TypeScript compilation failed."

# ------------------------------------------------------------------------------
# 3. Version Handling
# ------------------------------------------------------------------------------
if ($BumpType -eq "none") {
    $newVersion = $currentVersion
    Write-Host "`n==> 2. Keeping existing version ($newVersion)..." -ForegroundColor Yellow
} else {
    Write-Host "`n==> 2. Bumping Package Version ($BumpType)..." -ForegroundColor Cyan
    # Run npm version
    $newVersionRaw = & npm version $BumpType --no-git-tag-version
    if ($LASTEXITCODE -ne 0) {
         throw "Failed to bump package version."
    }
    $newVersion = $newVersionRaw.Trim()
    Write-Host "New Version: $newVersion" -ForegroundColor Green
}

# ------------------------------------------------------------------------------
# 4. Git Synchronization
# ------------------------------------------------------------------------------
Write-Host "`n==> 3. Checking Git Working Tree..." -ForegroundColor Cyan
Invoke-ExternalCommand -Command { git add . } -ErrorMessage "Failed to stage git files."

# Check if there are staged changes to commit
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
} else {
    Write-Host "No uncommitted file changes detected. Skipping commit step." -ForegroundColor Yellow
}

Write-Host "Pushing to remote repository (main)..." -ForegroundColor Gray
try {
    Invoke-ExternalCommand -Command { git push origin main } -ErrorMessage "Git push failed."
    Write-Host "GitHub synchronization complete!" -ForegroundColor Green
} catch {
    Write-Host "`n[ERROR] Git push failed. Please check your internet connection or git branch status." -ForegroundColor Red
    Write-Host "You can fix the issue and rerun .\publish.ps1 using option [0] to retry without bumping the version." -ForegroundColor Yellow
    exit 1
}

# ------------------------------------------------------------------------------
# 5. NPM Registry Publication
# ------------------------------------------------------------------------------
Write-Host "`n==> 4. NPM Registry Publication..." -ForegroundColor Cyan
$confirmPublish = Read-Host "Ready to publish $newVersion to NPM? [Y/n]"
if ($confirmPublish -eq "" -or $confirmPublish.ToUpper() -eq "Y") {
    try {
        Invoke-ExternalCommand -Command { npm publish --access public } -ErrorMessage "NPM publish failed."
        Write-Host "Successfully published $newVersion to NPM!" -ForegroundColor Green
    } catch {
        Write-Host "`n[ERROR] NPM publish failed." -ForegroundColor Red
        Write-Host "Verify your login status by running 'npm whoami' (must be logged in as 'bdking71')." -ForegroundColor Yellow
        Write-Host "To fix, run 'npm login' and then rerun .\publish.ps1 using option [0] to retry." -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "NPM Publish skipped by user. Local and Git changes remain intact." -ForegroundColor Yellow
}

Write-Host "`nAll operations completed successfully!" -ForegroundColor Green