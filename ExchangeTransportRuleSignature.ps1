# TelexPH - Automatic Email Signature via Exchange Transport Rules
# This adds signatures SERVER-SIDE to all outbound emails automatically

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
    DefaultPhone = "(044) 331 - 5040"
    DefaultAddress = "Cawayan Bugtong, Guimba, Nueva Ecija, Philippines"
    CompanyWebsite = "www.telexph.com"
    LogoUrl = "https://telexph.com/wp-content/uploads/2024/05/TELEX-Logo-2.png"
    DefaultPhotoUrl = "https://telexph.com/default-avatar.png"
}

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Get-SignatureHTML {
    param(
        [string]$DisplayName,
        [string]$JobTitle,
        [string]$Email,
        [string]$Phone,
        [string]$Address,
        [string]$PhotoUrl
    )
    
    $html = @"
<table cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; font-family: Arial, sans-serif; margin-top: 20px;">
    <tr>
        <td style="padding: 20px 0; border-top: 2px solid #8B1538;">
            <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td style="vertical-align: top; padding-right: 20px;">
                        <img src="$PhotoUrl" width="120" height="120" style="border-radius: 50%; border: 4px solid #8B1538; display: block;" alt="$DisplayName" />
                    </td>
                    <td style="border-left: 3px solid #8B1538; padding-left: 20px; vertical-align: top;">
                        <h2 style="margin: 0; font-size: 24px; color: #000000; font-style: italic;">$DisplayName</h2>
                        <div style="background-color: #8B1538; color: #ffffff; padding: 4px 12px; display: inline-block; font-size: 11px; font-weight: bold; margin-top: 5px; border-radius: 3px;">
                            $JobTitle
                        </div>
                        <p style="margin: 12px 0 0 0; font-size: 12px; color: #333333; line-height: 1.6;">
                            $Address<br>
                            $Phone<br>
                            <a href="mailto:$Email" style="color: #333333; text-decoration: none;">$Email</a><br>
                            <a href="https://$($config.CompanyWebsite)" style="color: #333333; text-decoration: none;">$($config.CompanyWebsite)</a>
                        </p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
"@
    return $html
}

function Get-UserPhotoUrl {
    param([string]$UserEmail)
    return $config.DefaultPhotoUrl
}

function Create-TransportRule {
    param(
        [string]$UserEmail,
        [string]$UserDisplayName,
        [string]$SignatureHTML
    )
    
    $ruleName = "Email Signature - $UserDisplayName"
    
    try {
        $existingRule = Get-TransportRule -Identity $ruleName -ErrorAction SilentlyContinue
        
        if ($existingRule) {
            Write-ColorOutput "   Rule already exists, updating..." "Yellow"
            
            Set-TransportRule -Identity $ruleName `
                -ApplyHtmlDisclaimerLocation 'Append' `
                -ApplyHtmlDisclaimerText $SignatureHTML `
                -ApplyHtmlDisclaimerFallbackAction 'Wrap' `
                -ErrorAction Stop
            
            Write-ColorOutput "   [OK] Rule updated: $ruleName" "Green"
        } else {
            Write-ColorOutput "   Creating new rule..." "Cyan"
            
            New-TransportRule -Name $ruleName `
                -FromScope 'InOrganization' `
                -From $UserEmail `
                -ApplyHtmlDisclaimerLocation 'Append' `
                -ApplyHtmlDisclaimerText $SignatureHTML `
                -ApplyHtmlDisclaimerFallbackAction 'Wrap' `
                -Mode 'Enforce' `
                -ErrorAction Stop
            
            Write-ColorOutput "   [OK] Rule created: $ruleName" "Green"
        }
        
        return $true
        
    } catch {
        Write-ColorOutput "   [ERROR] Failed: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Process-SingleUser {
    param([string]$Email)
    
    Write-ColorOutput "`n[PROCESSING] $Email" "Yellow"
    
    try {
        Write-ColorOutput "   Fetching user information..." "Cyan"
        $user = Get-User -Identity $Email -ErrorAction Stop
        $mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
        
        $displayName = $user.DisplayName
        $jobTitle = if ($mailbox.Title) { $mailbox.Title } else { "Team Member" }
        $phone = if ($mailbox.Phone) { $mailbox.Phone } else { $config.DefaultPhone }
        $address = if ($user.City) { "$($user.City), $($user.StateOrProvince)" } else { $config.DefaultAddress }
        
        Write-ColorOutput "   Name: $displayName" "Gray"
        Write-ColorOutput "   Title: $jobTitle" "Gray"
        
        $photoUrl = Get-UserPhotoUrl -UserEmail $Email
        
        Write-ColorOutput "   Generating signature..." "Cyan"
        $signatureHTML = Get-SignatureHTML `
            -DisplayName $displayName `
            -JobTitle $jobTitle `
            -Email $Email `
            -Phone $phone `
            -Address $address `
            -PhotoUrl $photoUrl
        
        $success = Create-TransportRule `
            -UserEmail $Email `
            -UserDisplayName $displayName `
            -SignatureHTML $signatureHTML
        
        return @{
            Success = $success
            Email = $Email
            Name = $displayName
        }
        
    } catch {
        Write-ColorOutput "   [ERROR] $($_.Exception.Message)" "Red"
        return @{
            Success = $false
            Email = $Email
            Error = $_.Exception.Message
        }
    }
}

function Process-MultipleUsers {
    param([array]$UserEmails)
    
    Write-ColorOutput "`n[BATCH] Processing $($UserEmails.Count) users...`n" "Green"
    
    $results = @()
    $counter = 0
    
    foreach ($email in $UserEmails) {
        $counter++
        Write-ColorOutput "[$counter/$($UserEmails.Count)]" "Gray"
        
        $result = Process-SingleUser -Email $email
        $results += $result
        
        Start-Sleep -Seconds 1
    }
    
    Write-ColorOutput "`n============================================================" "White"
    Write-ColorOutput "DEPLOYMENT SUMMARY" "Green"
    Write-ColorOutput "============================================================" "White"
    
    $successful = ($results | Where-Object { $_.Success }).Count
    $failed = ($results | Where-Object { -not $_.Success }).Count
    
    Write-ColorOutput "[OK] Successful: $successful" "Green"
    Write-ColorOutput "[ERROR] Failed: $failed" "Red"
    Write-ColorOutput "Total: $($results.Count)" "Cyan"
    
    if ($failed -gt 0) {
        Write-ColorOutput "`nFailed users:" "Red"
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-ColorOutput "   - $($_.Email): $($_.Error)" "Red"
        }
    }
    
    Write-ColorOutput "`nNEXT STEPS:" "Yellow"
    Write-ColorOutput "   1. Test by sending INTERNAL email (to @telexph.com)" "White"
    Write-ColorOutput "   2. Test by sending EXTERNAL email (to Gmail/Yahoo)" "White"
    Write-ColorOutput "   3. Signature appears on ALL emails (internal & external)" "White"
    Write-ColorOutput "   4. Signature WON'T appear in Outlook composer (normal)" "White"
    Write-ColorOutput "   5. Signature WILL appear in sent/received emails" "White"
    
    return $results
}

function Show-Usage {
    Write-ColorOutput "`nUsage:" "Yellow"
    Write-ColorOutput "  .\ExchangeTransportRuleSignature.ps1 -UserEmail user@telexph.com" "White"
    Write-ColorOutput "  .\ExchangeTransportRuleSignature.ps1 -BatchFile users.txt" "White"
    Write-ColorOutput "  .\ExchangeTransportRuleSignature.ps1 -AllUsers" "White"
}

# Main execution
try {
    Write-ColorOutput "`n============================================================" "Cyan"
    Write-ColorOutput "  TelexPH - Email Signature Automation" "Cyan"
    Write-ColorOutput "  Exchange Transport Rules (Automatic)" "Cyan"
    Write-ColorOutput "============================================================`n" "Cyan"
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-ColorOutput "[OK] Connected to Exchange Online" "Green"
    } catch {
        Write-ColorOutput "[ERROR] Not connected to Exchange Online!" "Red"
        Write-ColorOutput "Run: Connect-ExchangeOnline -UserPrincipalName innovation@telexph.com" "Yellow"
        exit 1
    }
    
    if ($UserEmail) {
        $result = Process-SingleUser -Email $UserEmail
        
        if ($result.Success) {
            Write-ColorOutput "`n[SUCCESS] Signature deployed for $($result.Name)" "Green"
            Write-ColorOutput "`nTEST NOW:" "Yellow"
            Write-ColorOutput "   1. Send email from $UserEmail to ANY address" "White"
            Write-ColorOutput "   2. Works for BOTH internal (@telexph.com) and external emails" "White"
            Write-ColorOutput "   3. Check received email - signature should be at bottom" "White"
        }
        
    } elseif ($BatchFile) {
        if (-not (Test-Path $BatchFile)) {
            Write-ColorOutput "[ERROR] File not found: $BatchFile" "Red"
            exit 1
        }
        
        $userEmails = Get-Content $BatchFile | 
            Where-Object { $_ -match '@' -and $_ -notmatch '^#' } |
            ForEach-Object { $_.Trim() }
        
        if ($userEmails.Count -eq 0) {
            Write-ColorOutput "[ERROR] No valid email addresses found in $BatchFile" "Red"
            exit 1
        }
        
        Process-MultipleUsers -UserEmails $userEmails
        
    } elseif ($AllUsers) {
        Write-ColorOutput "Fetching all users from organization..." "Cyan"
        $users = Get-Mailbox -ResultSize Unlimited | 
            Select-Object -ExpandProperty PrimarySmtpAddress
        
        Write-ColorOutput "Found $($users.Count) users" "Cyan"
        
        Process-MultipleUsers -UserEmails $users
        
    } else {
        Show-Usage
    }
    
} catch {
    Write-ColorOutput "`n[FATAL ERROR] $($_.Exception.Message)" "Red"
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}