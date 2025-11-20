# ============================
# 📧 Content Search & Purge Script
# ============================
# Version: 3.0.0
# Author: Mohsen Heidari
# Notes:
#   - Runs in both Windows PowerShell 5.1 and PowerShell 7+.
#   - Keeps emojis; save as UTF-8 with BOM for best display in 5.1.
#   - Auto-installs/updates ExchangeOnlineManagement:
#       • Admin => installs for AllUsers
#       • Non-admin => installs for CurrentUser
#   - Adaptive sign-in: prefers -DisableWAM, uses -EnableSearchOnlySession if available,
#     and falls back to credential prompt if interactive login fails.
# ============================

# ---------- [0] Environment & pre-flight ----------
# TLS 1.2 for secure operations
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

# Process-only policy to avoid global changes
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch { }

# Optional: improve emoji rendering (still save file as UTF-8 with BOM)
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch { }

# Version check
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "❌ PowerShell 5.1+ is required. Current: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    exit 1
}

# Admin detection (for global installs)
function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}
$IsAdmin = Test-IsAdmin

# ---------- [1] Banner ----------
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " 📧 Content Search Script - Automated Purge Tool " -ForegroundColor Cyan
Write-Host " ✅ Version 3.0.0 | Script by Abarkakia@gmail.com " -ForegroundColor Cyan
Write-Host " 🔍 Search by sender and subject in the last 7 days " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️ You will now be prompted to sign in. Choose your Purview-enabled admin account." -ForegroundColor Yellow
Write-Host ("ℹ️ PowerShell: {0} | Edition: {1} | Admin: {2}" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, $IsAdmin) -ForegroundColor DarkGray
Start-Sleep -Seconds 1

# ---------- [2] Module prerequisites (auto install/update) ----------
$requiredModule = "ExchangeOnlineManagement"
$minVersion     = [Version]"3.0.0"

try {
    $module = Get-Module -ListAvailable -Name $requiredModule | Sort-Object Version -Descending | Select-Object -First 1

    # Trust PSGallery temporarily (runtime only)
    try {
        $gallery = Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue
        if ($gallery -and $gallery.InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        }
    } catch { }

    if ($null -eq $module -or $module.Version -lt $minVersion) {
        Write-Host ("⬇️ Installing/Updating {0} (Current: {1})..." -f $requiredModule, ($module.Version)) -ForegroundColor Yellow
        if ($IsAdmin) {
            Install-Module $requiredModule -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
        } else {
            Write-Host "🔒 Not admin: installing module for CurrentUser." -ForegroundColor Yellow
            Install-Module $requiredModule -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }

    Import-Module $requiredModule -MinimumVersion $minVersion -ErrorAction Stop
    Write-Host "✅ ExchangeOnlineManagement loaded (Version: $((Get-Module $requiredModule).Version))" -ForegroundColor Green
}
catch {
    Write-Host "❌ Failed to install/import ${requiredModule}: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ---------- [3] Adaptive connect to Microsoft Purview (SCC) ----------
function Connect-ToPurview {
    # Avoid reusing a wrong cached identity from a previous session
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }

    # Detect available parameters for Connect-IPPSSession in *this* environment
    try {
        $cmd = Get-Command Connect-IPPSSession -ErrorAction Stop
        $availableParams = $cmd.Parameters.Keys
        Write-Host "🔍 Connect-IPPSSession parameters: $($availableParams -join ', ')" -ForegroundColor DarkGray
    }
    catch {
        throw "❌ Connect-IPPSSession cmdlet not found. Ensure ExchangeOnlineManagement is loaded."
    }

    # Preferred path: system browser (disable WAM) + search-only session if both available
    try {
        if (($availableParams -contains 'DisableWAM') -and ($availableParams -contains 'EnableSearchOnlySession')) {
            Write-Host "🔗 Using system browser sign-in (account picker) with Search-Only session..." -ForegroundColor Yellow
            Connect-IPPSSession -DisableWAM -EnableSearchOnlySession -ErrorAction Stop
        }
        elseif ($availableParams -contains 'DisableWAM') {
            Write-Host "🔗 Using system browser sign-in (account picker)..." -ForegroundColor Yellow
            Connect-IPPSSession -DisableWAM -ErrorAction Stop
        }
        elseif ($availableParams -contains 'EnableSearchOnlySession') {
            Write-Host "🔗 Using interactive sign-in (embedded) with Search-Only session..." -ForegroundColor Yellow
            Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop
        }
        else {
            Write-Host "🔗 Using interactive sign-in (embedded)..." -ForegroundColor Yellow
            Connect-IPPSSession -ErrorAction Stop
        }

        Write-Host "✅ Connected to Microsoft Purview (Security & Compliance)." -ForegroundColor Green
        return
    }
    catch {
        Write-Host "❌ Primary connect attempt failed: $($_.Exception.Message)" -ForegroundColor Red

        # Fallback: credential prompt (UPN+password); not ideal, but works in hardened environments
        if ($availableParams -contains 'Credential') {
            Write-Host "🔐 Trying secure credential prompt..." -ForegroundColor Yellow
            try {
                $cred = Get-Credential -Message "Enter Purview admin credentials"
                if ($availableParams -contains 'EnableSearchOnlySession') {
                    Connect-IPPSSession -Credential $cred -EnableSearchOnlySession -ErrorAction Stop
                } else {
                    Connect-IPPSSession -Credential $cred -ErrorAction Stop
                }
                Write-Host "✅ Connected to Microsoft Purview using credentials." -ForegroundColor Green
                return
            }
            catch {
                Write-Host "❌ Credential-based connect failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        Write-Host "💡 Troubleshooting tips for sign-in failures (e.g., 0xffffffff80070520):" -ForegroundColor Yellow
        Write-Host "   • Ensure Microsoft Edge WebView2 Runtime (Evergreen) is installed/updated on this host." -ForegroundColor Yellow
        Write-Host "   • Close all PowerShell windows and retry (pwsh -NoProfile is best)." -ForegroundColor Yellow
        Write-Host "   • If policies block embedded flows, use -DisableWAM to force system-browser flow." -ForegroundColor Yellow
        throw
    }
}
try { Connect-ToPurview } catch { Write-Host $_ -ForegroundColor Red; exit 1 }

# ---------- [4] Role check (best-effort) ----------
try {
    $assignments = Get-ManagementRoleAssignment -Role "Search And Purge" -ErrorAction Stop
    if (-not $assignments) {
        Write-Host "⚠️ Your account may not have the 'Search And Purge' role. Purge could fail." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "ℹ️ Skipping role check (role cmdlets may be limited in this session)." -ForegroundColor Yellow
}

# ---------- [5] Inputs ----------
Write-Host "This script creates a content search in Exchange Online and can purge the results." -ForegroundColor Red
Write-Host "*********************************************" -ForegroundColor Yellow
$senderEmail  = Read-Host "📨 Enter the sender's email address"
Write-Host "=============================================" -ForegroundColor Green
$emailSubject = Read-Host "📝 Enter the email subject (Press Enter for no subject)"
Write-Host "=============================================" -ForegroundColor Green
$searchName   = Read-Host "🔍 Enter the search name or ticket number"
Write-Host "=============================================" -ForegroundColor Green
$purgeMode    = Read-Host "⚠️ Do you want a quick purge (Y) or review results first (N)?"
Write-Host "*********************************************" -ForegroundColor Yellow

if ([string]::IsNullOrWhiteSpace($senderEmail)) { Write-Host "❌ Sender email is required." -ForegroundColor Red; exit 1 }
if ([string]::IsNullOrWhiteSpace($searchName)) { Write-Host "❌ Search name cannot be empty." -ForegroundColor Red; exit 1 }

# ---------- [6] Build KQL query for last 7 days ----------
$startDate      = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
$escapedSubject = $emailSubject -replace '"','\"'
$query          = "from:$senderEmail"
if (-not [string]::IsNullOrWhiteSpace($escapedSubject)) { $query += " AND subject:`"$escapedSubject`"" }
$query += " AND received>=$startDate"

Write-Host "🔎 Creating compliance search with query: $query"

# ---------- [7] Ensure unique name, create & start ----------
try {
    while (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
        Write-Host "⚠️ A compliance search named '$searchName' already exists." -ForegroundColor Yellow
        $searchName = Read-Host "🔁 Please enter a new, unique search name"
        if ([string]::IsNullOrWhiteSpace($searchName)) { Write-Host "❌ Search name cannot be empty." -ForegroundColor Red; exit 1 }
    }
} catch { }

try {
    New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $query -ErrorAction Stop | Out-Null
    Start-ComplianceSearch -Identity $searchName -ErrorAction Stop | Out-Null
    Write-Host "⏳ Search started. Waiting for results..." -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to create/start search: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ---------- [8] Spinner while waiting ----------
$spinner = @("😐","😑","😶","😮","😯","😲","😵","🤯")
$searchCompleted = $false
while (-not $searchCompleted) {
    try {
        $status = (Get-ComplianceSearch -Identity $searchName -ErrorAction Stop).Status
    } catch {
        $status = "Checking..."
    }
    if ($status -eq "Completed") { $searchCompleted = $true; break }
    foreach ($spinChar in $spinner) {
        Write-Host -NoNewline "`r$spinChar Waiting for search to complete... Current status: $status"
        Start-Sleep -Milliseconds 300
    }
}
Write-Host "`r✅ Search completed." -ForegroundColor Green

# ---------- [9] Results ----------
Write-Host "📦 Items found:"
try {
    Get-ComplianceSearch -Identity $searchName | Format-List Name,Status,Items
} catch {
    Write-Host "⚠️ Unable to display items count: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ---------- [10] Purge ----------
$doPurge = $false
if ($purgeMode -match '^(Y|y)$') { $doPurge = $true } else {
    $confirmPurge = Read-Host "❓ Do you want to purge the results? (Y/N)"
    if ($confirmPurge -match '^(Y|y)$') { $doPurge = $true }
}

if ($doPurge) {
    Write-Host "🚀 Initiating purge (SoftDelete)..." -ForegroundColor Yellow
    Start-Sleep -Seconds 30
    try {
        $action = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop
        Write-Host "🧹 Purge action submitted. Tracking..." -ForegroundColor Green

        $actionName = $action.Identity
        $terminal   = @("Completed","PartiallySucceeded","Failed","PartiallyFailed")
        do {
            Start-Sleep -Seconds 10
            try {
                $a = Get-ComplianceSearchAction -Identity $actionName -ErrorAction Stop
                Write-Host ("   • Purge status: {0}" -f $a.Status)
            } catch { $a = $null }
        } while ($a -and ($terminal -notcontains $a.Status))

        if ($a -and $a.Status -eq "Completed") { Write-Host "✅ Purge completed successfully." -ForegroundColor Green }
        elseif ($a) { Write-Host ("⚠️ Purge finished with status: {0}" -f $a.Status) -ForegroundColor Yellow }
        else { Write-Host "⚠️ Could not retrieve final purge action status." -ForegroundColor Yellow }
    } catch {
        Write-Host "❌ Purge failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "🛑 Purge cancelled by user."
}

# ---------- [11] Summary ----------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "✅ Script completed successfully." -ForegroundColor Green
Write-Host "📊 Search summary:" -ForegroundColor Yellow
Get-ComplianceSearch -Identity $searchName | Format-List Name,Status,Items
Write-Host "=============================================" -ForegroundColor Cyan

# ---------- [12] Cleanup ----------
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }