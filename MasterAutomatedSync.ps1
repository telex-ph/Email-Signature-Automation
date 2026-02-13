# TelexPH - MASTER AUTOMATED PHOTO SYNC
# Tries ALL methods to automatically fetch M365 profile photos
# NO MANUAL UPLOAD - FULLY AUTOMATED!

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Magenta
Write-Host "   MASTER AUTOMATED M365 PHOTO SYNC                           " -ForegroundColor Magenta
Write-Host "   Trying all methods for 100% automation!                    " -ForegroundColor Magenta
Write-Host "   NO MANUAL UPLOAD REQUIRED!                                 " -ForegroundColor Magenta
Write-Host "================================================================" -ForegroundColor Magenta
Write-Host ""

# Check Exchange connection
try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[OK] Connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] NOT CONNECTED TO EXCHANGE ONLINE!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
    Write-Host ""
    exit
}

# Get user info
try {
    $user = Get-User -Identity $UserEmail -ErrorAction Stop
    Write-Host "[USER] Found: $($user.DisplayName)" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "[ERROR] User not found: $UserEmail" -ForegroundColor Red
    exit
}

# Check current photo status
Write-Host "[CHECK] Checking current photo in Exchange mailbox..." -ForegroundColor Yellow
$mailbox = Get-Mailbox -Identity $UserEmail

if ($mailbox.ThumbnailPhoto -and $mailbox.ThumbnailPhoto.Length -gt 0) {
    $size = [math]::Round($mailbox.ThumbnailPhoto.Length / 1KB, 1)
    Write-Host "[RESULT] ✅ Photo already exists! Size: $size KB" -ForegroundColor Green
    Write-Host "[INFO] No sync needed - photo is already in Exchange!" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Deploy signature now:" -ForegroundColor Yellow
    Write-Host "  .\ExchangeTransportRuleSignature.ps1 -UserEmail $UserEmail" -ForegroundColor White
    Write-Host ""
    exit
}

Write-Host "[RESULT] No photo in Exchange mailbox" -ForegroundColor Red
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "   ATTEMPTING AUTOMATED SYNC (5 methods)                      " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$photoBytes = $null
$successMethod = ""

# ===================================================================
# METHOD 1: Set-UserPhoto -Save (Simplest - direct M365 sync)
# ===================================================================
Write-Host "[METHOD 1] Trying Set-UserPhoto -Save..." -ForegroundColor Yellow

try {
    Set-UserPhoto -Identity $UserEmail -Save -ErrorAction Stop
    Start-Sleep -Seconds 3
    
    $mailboxCheck = Get-Mailbox -Identity $UserEmail
    if ($mailboxCheck.ThumbnailPhoto -and $mailboxCheck.ThumbnailPhoto.Length -gt 0) {
        $size = [math]::Round($mailboxCheck.ThumbnailPhoto.Length / 1KB, 1)
        Write-Host "[SUCCESS] ✅ Method 1 worked! Photo synced ($size KB)" -ForegroundColor Green
        $successMethod = "Set-UserPhoto -Save"
        $photoBytes = $mailboxCheck.ThumbnailPhoto
    }
} catch {
    Write-Host "[FAILED] Method 1: $($_.Exception.Message)" -ForegroundColor DarkGray
}

# ===================================================================
# METHOD 2: Outlook REST API (HR size photos)
# ===================================================================
if (-not $photoBytes) {
    Write-Host ""
    Write-Host "[METHOD 2] Trying Outlook REST API..." -ForegroundColor Yellow
    
    $outlookUrls = @(
        "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=$UserEmail&size=HR648x648",
        "https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=$UserEmail&size=HR648x648"
    )
    
    foreach ($url in $outlookUrls) {
        try {
            Write-Host "[FETCH] Trying $url..." -ForegroundColor Gray
            $response = Invoke-WebRequest -Uri $url -Method Get -UseDefaultCredentials -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            
            if ($response.StatusCode -eq 200 -and $response.Content.Length -gt 1000) {
                $photoBytes = $response.Content
                $size = [math]::Round($photoBytes.Length / 1KB, 1)
                Write-Host "[SUCCESS] ✅ Method 2 worked! Downloaded photo ($size KB)" -ForegroundColor Green
                $successMethod = "Outlook REST API"
                break
            }
        } catch {
            Write-Host "[FAILED] $($_.Exception.Message)" -ForegroundColor DarkGray
        }
    }
}

# ===================================================================
# METHOD 3: Microsoft Graph REST API (via Azure CLI)
# ===================================================================
if (-not $photoBytes) {
    Write-Host ""
    Write-Host "[METHOD 3] Trying Microsoft Graph API via Azure CLI..." -ForegroundColor Yellow
    
    try {
        # Check if Azure CLI is available and logged in
        $account = az account show 2>$null | ConvertFrom-Json
        
        if ($account) {
            Write-Host "[AUTH] Azure CLI session found!" -ForegroundColor Green
            
            # Get access token for Microsoft Graph
            $tokenResult = az account get-access-token --resource https://graph.microsoft.com 2>$null | ConvertFrom-Json
            
            if ($tokenResult -and $tokenResult.accessToken) {
                Write-Host "[AUTH] Got Graph API access token" -ForegroundColor Green
                
                # Fetch photo from Graph API
                $headers = @{
                    "Authorization" = "Bearer $($tokenResult.accessToken)"
                }
                
                $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserEmail/photo/`$value"
                Write-Host "[FETCH] Downloading from: $graphUrl" -ForegroundColor Gray
                
                $photoBytes = Invoke-RestMethod -Uri $graphUrl -Headers $headers -Method Get -ErrorAction Stop
                
                if ($photoBytes -and $photoBytes.Length -gt 1000) {
                    $size = [math]::Round($photoBytes.Length / 1KB, 1)
                    Write-Host "[SUCCESS] ✅ Method 3 worked! Downloaded from Graph API ($size KB)" -ForegroundColor Green
                    $successMethod = "Microsoft Graph REST API"
                }
            }
        } else {
            Write-Host "[SKIP] Azure CLI not logged in" -ForegroundColor Gray
        }
    } catch {
        Write-Host "[FAILED] Method 3: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

# ===================================================================
# METHOD 4: Microsoft Graph PowerShell (if module available)
# ===================================================================
if (-not $photoBytes) {
    Write-Host ""
    Write-Host "[METHOD 4] Trying Microsoft Graph PowerShell..." -ForegroundColor Yellow
    
    try {
        # Check if Graph module is available
        $graphModule = Get-Module Microsoft.Graph.Users -ListAvailable -ErrorAction SilentlyContinue
        
        if ($graphModule) {
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            
            # Check if connected
            $context = Get-MgContext -ErrorAction SilentlyContinue
            
            if ($context) {
                Write-Host "[AUTH] Microsoft Graph connected!" -ForegroundColor Green
                
                $photoBytes = Get-MgUserPhoto -UserId $UserEmail -ErrorAction Stop
                
                if ($photoBytes -and $photoBytes.Length -gt 1000) {
                    $size = [math]::Round($photoBytes.Length / 1KB, 1)
                    Write-Host "[SUCCESS] ✅ Method 4 worked! Got photo from Graph PowerShell ($size KB)" -ForegroundColor Green
                    $successMethod = "Microsoft Graph PowerShell"
                }
            } else {
                Write-Host "[SKIP] Microsoft Graph not connected" -ForegroundColor Gray
            }
        } else {
            Write-Host "[SKIP] Microsoft Graph module not installed" -ForegroundColor Gray
        }
    } catch {
        Write-Host "[FAILED] Method 4: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

# ===================================================================
# METHOD 5: Get-UserPhoto cmdlet (if available)
# ===================================================================
if (-not $photoBytes) {
    Write-Host ""
    Write-Host "[METHOD 5] Trying Get-UserPhoto cmdlet..." -ForegroundColor Yellow
    
    try {
        $photo = Get-UserPhoto -Identity $UserEmail -ErrorAction Stop
        
        if ($photo -and $photo.PictureData -and $photo.PictureData.Length -gt 0) {
            $photoBytes = $photo.PictureData
            $size = [math]::Round($photoBytes.Length / 1KB, 1)
            Write-Host "[SUCCESS] ✅ Method 5 worked! Got photo from Get-UserPhoto ($size KB)" -ForegroundColor Green
            $successMethod = "Get-UserPhoto cmdlet"
        }
    } catch {
        Write-Host "[FAILED] Method 5: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

# ===================================================================
# UPLOAD TO EXCHANGE (if any method succeeded)
# ===================================================================
if ($photoBytes -and $photoBytes.Length -gt 0) {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "   UPLOADING TO EXCHANGE MAILBOX                              " -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "[INFO] Successfully fetched photo using: $successMethod" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "[UPLOAD] Uploading to Exchange mailbox..." -ForegroundColor Yellow
        
        Set-UserPhoto -Identity $UserEmail -PictureData $photoBytes -Confirm:$false -ErrorAction Stop
        
        # Verify upload
        $mailboxFinal = Get-Mailbox -Identity $UserEmail
        if ($mailboxFinal.ThumbnailPhoto -and $mailboxFinal.ThumbnailPhoto.Length -gt 0) {
            $finalSize = [math]::Round($mailboxFinal.ThumbnailPhoto.Length / 1KB, 1)
            
            Write-Host ""
            Write-Host "================================================================" -ForegroundColor Green
            Write-Host "   ✅ ✅ ✅ SUCCESS! AUTOMATED SYNC COMPLETE! ✅ ✅ ✅         " -ForegroundColor Green
            Write-Host "================================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "[METHOD USED] $successMethod" -ForegroundColor Cyan
            Write-Host "[PHOTO SIZE] $finalSize KB" -ForegroundColor Cyan
            Write-Host "[USER] $($user.DisplayName)" -ForegroundColor Cyan
            Write-Host "[EMAIL] $UserEmail" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "================================================================" -ForegroundColor Magenta
            Write-Host "   NEXT STEP: Deploy Email Signature                          " -ForegroundColor Magenta
            Write-Host "================================================================" -ForegroundColor Magenta
            Write-Host ""
            Write-Host "Run this command:" -ForegroundColor Yellow
            Write-Host "  .\ExchangeTransportRuleSignature.ps1 -UserEmail $UserEmail" -ForegroundColor White
            Write-Host ""
            Write-Host "🎉 NO MANUAL UPLOAD NEEDED! FULLY AUTOMATED! 🎉" -ForegroundColor Green
            Write-Host ""
            
        } else {
            Write-Host "[WARNING] Upload completed but verification failed" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "[ERROR] Failed to upload: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} else {
    # ALL METHODS FAILED
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host "   ❌ ALL AUTOMATED METHODS FAILED                            " -ForegroundColor Red
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "[INFO] Could not automatically fetch M365 profile photo." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "POSSIBLE REASONS:" -ForegroundColor Yellow
    Write-Host "  ❌ User has not uploaded photo to M365 profile" -ForegroundColor White
    Write-Host "  ❌ Photo sync is disabled in M365" -ForegroundColor White
    Write-Host "  ❌ No authentication method available (Azure CLI, Graph)" -ForegroundColor White
    Write-Host "  ❌ Network/firewall blocking photo access" -ForegroundColor White
    Write-Host ""
    Write-Host "TO ENABLE AUTOMATION:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1️⃣  SETUP AZURE CLI (RECOMMENDED - Enables full automation)" -ForegroundColor Yellow
    Write-Host "    Run: .\FixAzureCLIAuth.ps1" -ForegroundColor White
    Write-Host "    Then run this script again" -ForegroundColor White
    Write-Host "    This enables automation for ALL users!" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2️⃣  VERIFY USER HAS M365 PHOTO" -ForegroundColor Yellow
    Write-Host "    Ask $($user.DisplayName) to check:" -ForegroundColor White
    Write-Host "    https://myaccount.microsoft.com" -ForegroundColor Gray
    Write-Host "    If no photo → user needs to upload one first" -ForegroundColor Gray
    Write-Host ""
    Write-Host "3️⃣  ONE-TIME MANUAL (Not recommended for automation)" -ForegroundColor Yellow
    Write-Host "    Run: .\EasyPhotoUpload.ps1 -UserEmail $UserEmail" -ForegroundColor White
    Write-Host "    Only works for this one user" -ForegroundColor Gray
    Write-Host ""
    Write-Host "RECOMMENDED: Setup Azure CLI for permanent automation! ✅" -ForegroundColor Cyan
    Write-Host ""
}