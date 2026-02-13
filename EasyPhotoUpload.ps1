# TelexPH - Easy Photo Upload Tool
# Interactive file browser for uploading profile photos

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Magenta
Write-Host "   TelexPH - Easy Photo Upload Tool                         " -ForegroundColor Magenta
Write-Host "   Upload profile photo to Exchange Online mailbox          " -ForegroundColor Magenta
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

# Check current photo
$mailbox = Get-Mailbox -Identity $UserEmail
if ($mailbox.ThumbnailPhoto -and $mailbox.ThumbnailPhoto.Length -gt 0) {
    $currentSize = [math]::Round($mailbox.ThumbnailPhoto.Length / 1KB, 1)
    Write-Host "[INFO] Current photo exists in mailbox ($currentSize KB)" -ForegroundColor Cyan
    Write-Host ""
}

# Open file browser
Write-Host "[SELECT] Please select a photo file..." -ForegroundColor Yellow
Write-Host ""

Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Select Profile Photo"
$openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.gif)|*.jpg;*.jpeg;*.png;*.gif|All Files (*.*)|*.*"
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

$result = $openFileDialog.ShowDialog()

if ($result -eq 'OK') {
    $photoPath = $openFileDialog.FileName
    Write-Host "[SELECTED] $photoPath" -ForegroundColor Green
    
    # Check file size
    $fileInfo = Get-Item $photoPath
    $fileSizeKB = [math]::Round($fileInfo.Length / 1KB, 1)
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    
    Write-Host "[SIZE] File size: $fileSizeKB KB ($fileSizeMB MB)" -ForegroundColor Gray
    
    if ($fileInfo.Length -gt 10MB) {
        Write-Host "[WARNING] File is very large! Microsoft 365 recommends photos under 100 KB." -ForegroundColor Yellow
        Write-Host "[INFO] Large photos may cause issues with email clients." -ForegroundColor Yellow
        Write-Host ""
        $continue = Read-Host "Continue anyway? (Y/N)"
        if ($continue -ne "Y" -and $continue -ne "y") {
            Write-Host "[CANCELLED] Photo upload cancelled" -ForegroundColor Yellow
            exit
        }
    }
    
    Write-Host ""
    Write-Host "[UPLOAD] Uploading photo to Exchange mailbox..." -ForegroundColor Yellow
    
    try {
        # Read photo file
        $photoBytes = [System.IO.File]::ReadAllBytes($photoPath)
        
        # Upload to Exchange Online
        Set-UserPhoto -Identity $UserEmail -PictureData $photoBytes -Confirm:$false -ErrorAction Stop
        
        Write-Host "[SUCCESS] Photo uploaded successfully!" -ForegroundColor Green
        Write-Host ""
        
        # Verify upload
        $updatedMailbox = Get-Mailbox -Identity $UserEmail
        if ($updatedMailbox.ThumbnailPhoto -and $updatedMailbox.ThumbnailPhoto.Length -gt 0) {
            $uploadedSize = [math]::Round($updatedMailbox.ThumbnailPhoto.Length / 1KB, 1)
            Write-Host "[VERIFY] Photo stored in mailbox ($uploadedSize KB)" -ForegroundColor Green
            
            Write-Host ""
            Write-Host "================================================================" -ForegroundColor Green
            Write-Host "   SUCCESS! Profile photo uploaded to Exchange mailbox       " -ForegroundColor Green
            Write-Host "================================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "Next steps:" -ForegroundColor Yellow
            Write-Host "  1. Run signature script:" -ForegroundColor White
            Write-Host "     .\ExchangeTransportRuleSignature.ps1 -UserEmail $UserEmail" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  2. Send a test email to verify signature with photo" -ForegroundColor White
            Write-Host ""
            
        } else {
            Write-Host "[WARNING] Upload completed but photo not found in mailbox" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "[ERROR] Failed to upload photo: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        # Common error solutions
        if ($_.Exception.Message -like "*size*") {
            Write-Host "SOLUTION: Photo file is too large. Try these:" -ForegroundColor Yellow
            Write-Host "  1. Compress photo using: https://tinyjpg.com" -ForegroundColor White
            Write-Host "  2. Resize photo to 500x500 pixels or smaller" -ForegroundColor White
            Write-Host "  3. Keep file size under 100 KB" -ForegroundColor White
        } elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access*") {
            Write-Host "SOLUTION: Permission denied. Try these:" -ForegroundColor Yellow
            Write-Host "  1. Run PowerShell as Administrator" -ForegroundColor White
            Write-Host "  2. Reconnect to Exchange: Connect-ExchangeOnline" -ForegroundColor White
            Write-Host "  3. Check if you have admin rights for this mailbox" -ForegroundColor White
        } else {
            Write-Host "SOLUTION: Try these troubleshooting steps:" -ForegroundColor Yellow
            Write-Host "  1. Check if file is corrupted (try opening in image viewer)" -ForegroundColor White
            Write-Host "  2. Convert to JPG format" -ForegroundColor White
            Write-Host "  3. Reconnect to Exchange Online" -ForegroundColor White
        }
    }
    
} else {
    Write-Host "[CANCELLED] No file selected" -ForegroundColor Yellow
}

Write-Host ""