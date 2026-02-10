# TelexPH Email Signature Deployment via Exchange Online
# This script sets email signatures using Exchange Online PowerShell
# More reliable than Graph API for signature management

#Requires -Modules ExchangeOnlineManagement

param(
    [Parameter(Mandatory=$false)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$BatchFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$AllUsers
)

# Configuration
$config = @{
    LogoUrl = "https://storage.googleapis.com/msgsndr/KlBL9XEG0eVNlAqE7m5V/media/69804d9df7a877373924ac5d.png"
    DefaultPhone = "(044) 331 - 5040"
    DefaultAddress = "Cawayan Bugtong,Guimba, Nueva Ecija, Philippines"
    DefaultPhotoUrl = "https://telexph.com/default-avatar.png"
    CompanyWebsite = "www.telexph.com"
    FacebookUrl = "https://www.facebook.com/telexph"
    InstagramUrl = "https://www.instagram.com/telexph"
    LinkedInUrl = "https://www.linkedin.com/company/telexph"
}

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-EmailSignatureHTML {
    param(
        [object]$User,
        [string]$PhotoUrl
    )
    
    $displayName = $User.DisplayName
    $jobTitle = if ($User.Title) { $User.Title } else { "Team Member" }
    $email = $User.PrimarySmtpAddress
    $phone = if ($User.MobilePhone) { $User.MobilePhone } else { $config.DefaultPhone }
    $address = if ($User.Office) { $User.Office } else { $config.DefaultAddress }
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; }
    </style>
</head>
<body>
    <table cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; font-family: 'Montserrat', Arial, sans-serif;">
        <tr>
            <td style="padding: 20px 0;">
                <table cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                        <!-- Profile Photo -->
                        <td style="width: 150px; vertical-align: top; padding-right: 20px;">
                            <img src="$PhotoUrl" alt="$displayName" style="width: 140px; height: 140px; border-radius: 50%; border: 4px solid #8B1538; object-fit: cover; display: block;" />
                        </td>
                        
                        <!-- Info Section -->
                        <td style="vertical-align: top; padding-right: 20px;">
                            <h1 style="margin: 0 0 10px 0; font-size: 28px; font-weight: 900; font-style: italic; color: #000000;">$displayName</h1>
                            
                            <div style="background-color: #8B1538; color: #ffffff; padding: 4px 12px; display: inline-block; font-size: 12px; font-weight: 600; margin-bottom: 15px; border-radius: 3px;">
                                $jobTitle
                            </div>
                            
                            <table cellpadding="0" cellspacing="0" border="0" style="margin-top: 10px;">
                                <tr>
                                    <td style="padding: 5px 0;">
                                        <img src="https://img.icons8.com/color/20/000000/marker.png" alt="Location" style="vertical-align: middle; margin-right: 8px;" />
                                        <span style="font-size: 13px; color: #333333; vertical-align: middle;">$address</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px 0;">
                                        <img src="https://img.icons8.com/color/20/000000/phone.png" alt="Phone" style="vertical-align: middle; margin-right: 8px;" />
                                        <span style="font-size: 13px; color: #333333; vertical-align: middle;">$phone</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px 0;">
                                        <img src="https://img.icons8.com/color/20/000000/email.png" alt="Email" style="vertical-align: middle; margin-right: 8px;" />
                                        <a href="mailto:$email" style="font-size: 13px; color: #333333; text-decoration: none; vertical-align: middle;">$email</a>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px 0;">
                                        <img src="https://img.icons8.com/color/20/000000/domain.png" alt="Website" style="vertical-align: middle; margin-right: 8px;" />
                                        <a href="https://$($config.CompanyWebsite)" target="_blank" style="font-size: 13px; color: #333333; text-decoration: none; vertical-align: middle;">$($config.CompanyWebsite)</a>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        
                        <!-- Logo and Social Media -->
                        <td style="width: 150px; vertical-align: top; text-align: center;">
                            <img src="https://telexph.com/logo.png" alt="TELEX Logo" style="width: 120px; margin-bottom: 20px; display: block; margin-left: auto; margin-right: auto;" />
                            
                            <div style="text-align: center;">
                                <a href="$($config.FacebookUrl)" target="_blank" style="text-decoration: none; margin: 0 5px;">
                                    <img src="https://img.icons8.com/color/30/000000/facebook.png" alt="Facebook" style="width: 28px; height: 28px;" />
                                </a>
                                <a href="$($config.InstagramUrl)" target="_blank" style="text-decoration: none; margin: 0 5px;">
                                    <img src="https://img.icons8.com/color/30/000000/instagram-new.png" alt="Instagram" style="width: 28px; height: 28px;" />
                                </a>
                                <a href="$($config.LinkedInUrl)" target="_blank" style="text-decoration: none; margin: 0 5px;">
                                    <img src="https://img.icons8.com/color/30/000000/linkedin.png" alt="LinkedIn" style="width: 28px; height: 28px;" />
                                </a>
                            </div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@
    
    return $html
}

function Set-UserEmailSignature {
    param(
        [string]$UserEmail,
        [string]$SignatureHTML
    )
    
    try {
        # Set mailbox message configuration
        Set-MailboxMessageConfiguration -Identity $UserEmail `
            -SignatureHTML $SignatureHTML `
            -AutoAddSignature $true `
            -AutoAddSignatureOnReply $true
        
        Write-ColorOutput "✅ Signature set successfully for $UserEmail" "Green"
        
        # Also save to file for backup
        $fileName = $UserEmail.Replace("@", "_at_")
        $outputPath = ".\signatures\$fileName.html"
        
        if (-not (Test-Path ".\signatures")) {
            New-Item -ItemType Directory -Path ".\signatures" | Out-Null
        }
        
        $SignatureHTML | Out-File -FilePath $outputPath -Encoding UTF8
        Write-ColorOutput "📄 Backup saved to: $outputPath" "Cyan"
        
        return $true
        
    } catch {
        Write-ColorOutput "❌ Error setting signature for ${UserEmail}: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Process-SingleUser {
    param(
        [string]$Email
    )
    
    Write-ColorOutput "`n🔄 Processing user: $Email" "Yellow"
    
    try {
        # Get user mailbox information
        Write-ColorOutput "📥 Fetching user information..." "Cyan"
        $mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
        
        # Get user photo (if available)
        Write-ColorOutput "📸 Fetching user photo..." "Cyan"
        try {
            $photo = Get-UserPhoto -Identity $Email -ErrorAction Stop
            $photoUrl = "data:image/jpeg;base64," + [Convert]::ToBase64String($photo.PictureData)
        } catch {
            Write-ColorOutput "⚠️  No photo found, using default" "Yellow"
            $photoUrl = $config.DefaultPhotoUrl
        }
        
        # Generate signature HTML
        Write-ColorOutput "📝 Generating signature..." "Cyan"
        $signatureHTML = Get-EmailSignatureHTML -User $mailbox -PhotoUrl $photoUrl
        
        # Set signature
        Write-ColorOutput "📤 Setting signature..." "Cyan"
        $result = Set-UserEmailSignature -UserEmail $Email -SignatureHTML $signatureHTML
        
        return @{
            Success = $result
            Email = $Email
            Name = $mailbox.DisplayName
        }
        
    } catch {
        Write-ColorOutput "❌ Failed to process ${Email}: $($_.Exception.Message)" "Red"
        return @{
            Success = $false
            Email = $Email
            Error = $_.Exception.Message
        }
    }
}

function Process-MultipleUsers {
    param(
        [array]$UserEmails
    )
    
    Write-ColorOutput "`n🚀 Starting batch processing for $($UserEmails.Count) users...`n" "Green"
    
    $results = @()
    $counter = 0
    
    foreach ($email in $UserEmails) {
        $counter++
        Write-ColorOutput "[$counter/$($UserEmails.Count)]" "Gray"
        
        $result = Process-SingleUser -Email $email
        $results += $result
        
        # Small delay to avoid throttling
        Start-Sleep -Milliseconds 500
    }
    
    # Summary
    Write-ColorOutput "`n$('=' * 50)" "White"
    Write-ColorOutput "📊 BATCH PROCESSING SUMMARY" "Green"
    Write-ColorOutput "$('=' * 50)" "White"
    
    $successful = ($results | Where-Object { $_.Success }).Count
    $failed = ($results | Where-Object { -not $_.Success }).Count
    
    Write-ColorOutput "✅ Successful: $successful" "Green"
    Write-ColorOutput "❌ Failed: $failed" "Red"
    Write-ColorOutput "📧 Total: $($results.Count)" "Cyan"
    
    if ($failed -gt 0) {
        Write-ColorOutput "`n❌ Failed users:" "Red"
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-ColorOutput "   - $($_.Email): $($_.Error)" "Red"
        }
    }
    
    return $results
}

# Main execution
try {
    # Check if Exchange Online module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-ColorOutput "❌ ExchangeOnlineManagement module not found!" "Red"
        Write-ColorOutput "Install with: Install-Module -Name ExchangeOnlineManagement" "Yellow"
        exit 1
    }
    
    # Connect to Exchange Online
    Write-ColorOutput "🔐 Connecting to Exchange Online..." "Cyan"
    Connect-ExchangeOnline -ShowBanner:$false
    
    # Process based on parameters
    if ($UserEmail) {
        # Single user mode
        Process-SingleUser -Email $UserEmail
        
    } elseif ($BatchFile) {
        # Batch mode from file
        if (-not (Test-Path $BatchFile)) {
            Write-ColorOutput "❌ File not found: $BatchFile" "Red"
            exit 1
        }
        
        $userEmails = Get-Content $BatchFile | Where-Object { $_ -match '@' }
        Process-MultipleUsers -UserEmails $userEmails
        
    } elseif ($AllUsers) {
        # Process all users
        Write-ColorOutput "🔄 Fetching all users from organization..." "Cyan"
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }
        $userEmails = $mailboxes | Select-Object -ExpandProperty PrimarySmtpAddress
        
        Process-MultipleUsers -UserEmails $userEmails
        
    } else {
        # Show usage
        Write-ColorOutput "`nUsage:" "Yellow"
        Write-ColorOutput "  .\DeployEmailSignatures.ps1 -UserEmail user@telexph.com" "White"
        Write-ColorOutput "  .\DeployEmailSignatures.ps1 -BatchFile users.txt" "White"
        Write-ColorOutput "  .\DeployEmailSignatures.ps1 -AllUsers" "White"
    }
    
} catch {
    Write-ColorOutput "`n❌ Fatal error: $($_.Exception.Message)" "Red"
    exit 1
} finally {
    # Disconnect from Exchange Online
    Write-ColorOutput "`n🔌 Disconnecting from Exchange Online..." "Cyan"
    Disconnect-ExchangeOnline -Confirm:$false
}