# ==============================================================================
# Automate Build, Git Push, and NPM Package Publication
# Usage: .\publish.ps1 -CommitMessage "Updates and fixes" -BumpType patch
# ==============================================================================
param (
    [string]$CommitMessage = "Update spsignature library and documentation",
    [ValidateSet("patch", "minor", "major")]
    [string]$BumpType = "patch"
)

$ErrorActionPreference = "Stop"

Write-Host "==> 1. Running TypeScript Compilation..." -ForegroundColor Cyan
npm run build

Write-Host "==> 2. Bumping Package Version ($BumpType)..." -ForegroundColor Cyan
$newVersion = npm version $BumpType --no-git-tag-version
Write-Host "New Version: $newVersion" -ForegroundColor Green

Write-Host "==> 3. Committing and Pushing to GitHub..." -ForegroundColor Cyan
git add .
git commit -m "$CommitMessage (v$newVersion)"
git push origin main
Write-Host "GitHub sync complete!" -ForegroundColor Green

Write-Host "==> 4. Publishing to NPM Registry..." -ForegroundColor Cyan
npm publish --access public
Write-Host "Successfully published $newVersion to NPM!" -ForegroundColor Green

Write-Host "`nAll operations completed successfully." -ForegroundColor Green