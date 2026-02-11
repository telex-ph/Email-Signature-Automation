# TelexPH Email Signature Deployment via Microsoft Graph API
# This script generates signatures and emails them to users
# NO ADMIN ROLE REQUIRED - uses Mail.ReadWrite and User.Read.All permissions

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Mail

param(
    [Parameter(Mandatory=$false)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$BatchFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$AllUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$SenderEmail = "hjreyes@telexph.com"  # Your email address
)

# Configuration
$config = @{
    DefaultPhone = "(044) 331 - 5040"
    DefaultAddress = "Cawayan Bugtong, Guimba, Nueva Ecija, Philippines"
    DefaultPhotoUrl = "https://telexph.com/default-avatar.png"
    CompanyWebsite = "www.telexph.com"
    LogoUrl = "https://i.imgur.com/YOUR_LOGO_HERE.png"
    FacebookUrl = "https://www.facebook.com/telexph"
    InstagramUrl = "https://www.instagram.com/telexph"
    LinkedInUrl = "https://www.linkedin.com/company/telexph"
}

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Get-EmailSignatureHTML {
    param(
        [object]$User,
        [string]$PhotoUrl
    )
    
    $displayName = $User.DisplayName
    $jobTitle = if ($User.JobTitle) { $User.JobTitle } else { "Team Member" }
    $email = $User.Mail
    $phone = if ($User.MobilePhone) { $User.MobilePhone } else { $config.DefaultPhone }
    $address = if ($User.OfficeLocation) { $User.OfficeLocation } else { $config.DefaultAddress }
    
    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; }
    </style>
</head>
<body>
    <table cellpadding="0" cellspacing="0" border="0" style="max-width: 700px; font-family: 'Montserrat', Arial, sans-serif;">
        <tr>
            <td style="padding: 20px 0;">
                <table cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                        <td style="width: 200px; vertical-align: top; padding-right: 0px;">
                            <img src="$PhotoUrl" alt="$displayName" style="width: 180px; height: 180px; border-radius: 50%; border: 6px solid #8B1538; object-fit: cover; display: block;" />
                        </td>
                        <td style="vertical-align: top; padding-right: 20px; border-left: 4px solid #8B1538; padding-left: 20px;">
                            <h1 style="margin: 0 0 12px 0; font-size: 36px; font-weight: 900; font-style: italic; color: #000000; line-height: 1; white-space: nowrap;">$displayName</h1>
                            <div style="background-color: #8B1538; color: #ffffff; padding: 6px 18px; display: inline-block; font-size: 13px; font-weight: 700; margin-bottom: 20px; border-radius: 5px;">$jobTitle</div>
                            <table cellpadding="0" cellspacing="0" border="0" style="margin-top: 10px;">
                                <tr><td style="padding: 6px 0;"><span style="font-size: 14px; color: #333333;">📍 $address</span></td></tr>
                                <tr><td style="padding: 6px 0;"><span style="font-size: 14px; color: #333333;">📞 $phone</span></td></tr>
                                <tr><td style="padding: 6px 0;"><a href="mailto:$email" style="font-size: 14px; color: #333333; text-decoration: none;">📧 $email</a></td></tr>
                                <tr><td style="padding: 6px 0;"><a href="https://$($config.CompanyWebsite)" target="_blank" style="font-size: 14px; color: #333333; text-decoration: none;">🌐 $($config.CompanyWebsite)</a></td></tr>
                            </table>
                        </td>
                        <td style="width: 200px; vertical-align: top; text-align: center; padding-left: 20px;">
                            <img src="$($config.LogoUrl)" alt="TELEX Logo" style="width: 150px; height: auto; margin-bottom: 25px; display: block; margin-left: auto; margin-right: auto;" />
                            <div style="text-align: center;">
                                <a href="$($config.FacebookUrl)" style="margin: 0 5px;">Facebook</a>
                                <a href="$($config.InstagramUrl)" style="margin: 0 5px;">Instagram</a>
                                <a href="$($config.LinkedInUrl)" style="margin: 0 5px;">LinkedIn</a>
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
}

function Get-InstructionEmailHTML {
    param(
        [string]$UserName,
        [string]$SignatureHTML
    )
    
    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 20px; }
        .header { background-color: #8B1538; color: white; padding: 20px; border-radius: 5px; }
        .content { background-color: #f9f9f9; padding: 20px; border-radius: 5px; margin: 20px 0; }
        .steps { background-color: white; padding: 15px; border-left: 4px solid #8B1538; margin: 10px 0; }
        .step { margin: 10px 0; padding-left: 20px; }
        .signature-preview { border: 2px solid #ddd; padding: 20px; margin: 20px 0; background-color: white; }
        .button { background-color: #8B1538; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; margin: 10px 0; }
        code { background-color: #f4f4f4; padding: 2px 5px; border-radius: 3px; font-family: monospace; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🎨 Your New TELEX Email Signature</h1>
        <p>Professional signature ready to install!</p>
    </div>
    
    <div class="content">
        <p>Hi <strong>$UserName</strong>,</p>
        <p>Your personalized email signature has been generated! Follow the simple steps below to install it in Outlook.</p>
    </div>
    
    <div class="steps">
        <h2>📋 Installation Instructions</h2>
        
        <div class="step">
            <strong>Step 1:</strong> Copy the signature HTML below (CTRL+A to select all, CTRL+C to copy)
        </div>
        
        <div class="step">
            <strong>Step 2:</strong> Open Outlook Desktop App or Outlook Web
        </div>
        
        <div class="step">
            <strong>For Outlook Desktop:</strong>
            <ul>
                <li>File → Options → Mail → Signatures</li>
                <li>Click "New" and name it (e.g., "TELEX Signature")</li>
                <li>Click in the editor area</li>
                <li>Press CTRL+V to paste</li>
                <li>Set as default for new messages and replies</li>
                <li>Click OK</li>
            </ul>
        </div>
        
        <div class="step">
            <strong>For Outlook Web (office.com):</strong>
            <ul>
                <li>Click Settings (⚙️) → View all Outlook settings</li>
                <li>Mail → Compose and reply → Email signature</li>
                <li>Paste the HTML in the editor (CTRL+V)</li>
                <li>Check "Automatically include my signature on new messages"</li>
                <li>Click Save</li>
            </ul>
        </div>
    </div>
    
    <div class="signature-preview">
        <h3>📧 Your Signature Preview:</h3>
        $SignatureHTML
    </div>
    
    <div class="content">
        <h3>📝 Copy the HTML below:</h3>
        <p><em>Select all the text in the box below and copy it (CTRL+A then CTRL+C):</em></p>
        <textarea style="width: 100%; height: 150px; font-family: monospace; padding: 10px; border: 1px solid #ddd; border-radius: 5px;">$($SignatureHTML -replace '<','&lt;' -replace '>','&gt;')</textarea>
    </div>
    
    <div class="content">
        <p><strong>Need help?</strong> Contact IT support or reply to this email.</p>
        <p>Best regards,<br>TELEX IT Team</p>
    </div>
</body>
</html>
"@
}

function Send-SignatureEmail {
    param(
        [string]$RecipientEmail,
        [string]$RecipientName,
        [string]$SignatureHTML,
        [string]$SenderEmail
    )
    
    try {
        $subject = "🎨 Your New TELEX Email Signature - Setup Required"
        $emailBody = Get-InstructionEmailHTML -UserName $RecipientName -SignatureHTML $SignatureHTML
        
        $message = @{
            subject = $subject
            body = @{
                contentType = "HTML"
                content = $emailBody
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $RecipientEmail
                    }
                }
            )
        }
        
        Send-MgUserMail -UserId $SenderEmail -Message $message -SaveToSentItems
        Write-ColorOutput "✅ Email sent to $RecipientEmail" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ Failed to send email to ${RecipientEmail}: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Process-SingleUser {
    param(
        [string]$Email,
        [string]$SenderEmail
    )
    
    Write-ColorOutput "`n🔄 Processing: $Email" "Yellow"
    
    try {
        Write-ColorOutput "📥 Fetching user information..." "Cyan"
        $user = Get-MgUser -UserId $Email -Property "DisplayName,Mail,JobTitle,MobilePhone,OfficeLocation" -ErrorAction Stop
        
        Write-ColorOutput "📸 Fetching user photo..." "Cyan"
        try {
            $photoStream = Get-MgUserPhotoContent -UserId $Email -ErrorAction Stop
            $photoBytes = $photoStream.ToArray()
            $photoBase64 = [Convert]::ToBase64String($photoBytes)
            $photoUrl = "data:image/jpeg;base64,$photoBase64"
        } catch {
            Write-ColorOutput "⚠️  No photo found, using default" "Yellow"
            $photoUrl = $config.DefaultPhotoUrl
        }
        
        Write-ColorOutput "📝 Generating signature..." "Cyan"
        $signatureHTML = Get-EmailSignatureHTML -User $user -PhotoUrl $photoUrl
        
        # Save to file
        $fileName = $Email.Replace("@", "_at_")
        $outputPath = ".\signatures\$fileName.html"
        
        if (-not (Test-Path ".\signatures")) {
            New-Item -ItemType Directory -Path ".\signatures" | Out-Null
        }
        
        $signatureHTML | Out-File -FilePath $outputPath -Encoding UTF8
        Write-ColorOutput "📄 Signature saved: $outputPath" "Cyan"
        
        Write-ColorOutput "📤 Sending email with instructions..." "Cyan"
        $emailSent = Send-SignatureEmail -RecipientEmail $Email -RecipientName $user.DisplayName -SignatureHTML $signatureHTML -SenderEmail $SenderEmail
        
        return @{
            Success = $emailSent
            Email = $Email
            Name = $user.DisplayName
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
        [array]$UserEmails,
        [string]$SenderEmail
    )
    
    Write-ColorOutput "`n🚀 Processing $($UserEmails.Count) users...`n" "Green"
    
    $results = @()
    $counter = 0
    
    foreach ($email in $UserEmails) {
        $counter++
        Write-ColorOutput "[$counter/$($UserEmails.Count)]" "Gray"
        
        $result = Process-SingleUser -Email $email -SenderEmail $SenderEmail
        $results += $result
        
        Start-Sleep -Seconds 2  # Rate limiting
    }
    
    Write-ColorOutput "`n================================================" "White"
    Write-ColorOutput "📊 SUMMARY" "Green"
    $successful = ($results | Where-Object { $_.Success }).Count
    $failed = ($results | Where-Object { -not $_.Success }).Count
    Write-ColorOutput "✅ Successful: $successful" "Green"
    Write-ColorOutput "❌ Failed: $failed" "Red"
    
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
    # Check for required modules
    $requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Mail")
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-ColorOutput "❌ Module $module not found!" "Red"
            Write-ColorOutput "Install with: Install-Module -Name $module -Scope CurrentUser" "Yellow"
            exit 1
        }
    }
    
    Write-ColorOutput "🔐 Connecting to Microsoft Graph..." "Cyan"
    Connect-MgGraph -Scopes "User.Read.All", "Mail.Send" -NoWelcome
    
    if ($UserEmail) {
        Process-SingleUser -Email $UserEmail -SenderEmail $SenderEmail
        
    } elseif ($BatchFile) {
        if (-not (Test-Path $BatchFile)) {
            Write-ColorOutput "❌ File not found: $BatchFile" "Red"
            exit 1
        }
        
        $userEmails = Get-Content $BatchFile | Where-Object { $_ -match '@' -and $_ -notmatch '^#' }
        Process-MultipleUsers -UserEmails $userEmails -SenderEmail $SenderEmail
        
    } elseif ($AllUsers) {
        Write-ColorOutput "🔄 Fetching all users..." "Cyan"
        $users = Get-MgUser -All -Filter "accountEnabled eq true" -Property "Mail"
        $userEmails = $users | Where-Object { $_.Mail } | Select-Object -ExpandProperty Mail
        
        Process-MultipleUsers -UserEmails $userEmails -SenderEmail $SenderEmail
        
    } else {
        Write-ColorOutput "`nUsage:" "Yellow"
        Write-ColorOutput "  .\DeployEmailSignatures_GraphAPI.ps1 -UserEmail user@telexph.com" "White"
        Write-ColorOutput "  .\DeployEmailSignatures_GraphAPI.ps1 -BatchFile users.txt" "White"
        Write-ColorOutput "  .\DeployEmailSignatures_GraphAPI.ps1 -AllUsers" "White"
        Write-ColorOutput "  .\DeployEmailSignatures_GraphAPI.ps1 -BatchFile users.txt -SenderEmail admin@telexph.com" "White"
    }
    
} catch {
    Write-ColorOutput "`n❌ Fatal error: $($_.Exception.Message)" "Red"
    exit 1
    
} finally {
    Write-ColorOutput "`n🔌 Disconnecting from Microsoft Graph..." "Cyan"
    Disconnect-MgGraph
    Write-ColorOutput "✅ Done!" "Green"
}