# TelexPH - Department Email Signature
# - NO profile photo, NO job title
# - Department name as header instead
# - Same responsive layout as personal signature
# - Background image support
# - Mobile-friendly (no min-width restriction)

#Requires -Modules ExchangeOnlineManagement

param(
    [Parameter(Mandatory=$false)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$BatchFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$AllUsers,
    
    [Parameter(Mandatory=$true)]
    [string]$CloudinaryCloudName,
    
    [Parameter(Mandatory=$true)]
    [string]$CloudinaryApiKey,
    
    [Parameter(Mandatory=$true)]
    [string]$CloudinaryApiSecret,
    
    [Parameter(Mandatory=$false)]
    [string]$CloudinaryFolder = "email-signatures",
    
    [Parameter(Mandatory=$false)]
    [string]$BackgroundImageUrl = ""
)

# --- Configuration ---
$config = @{
    DefaultPhone    = "(044) 331 - 5040"
    DefaultAddress  = "Nueva Ecija, Philippines"
    CompanyWebsite  = "www.telexph.com"
    LogoUrl         = "https://assets.cdn.filesafe.space/KlBL9XEG0eVNlAqE7m5V/media/69981f9ef83453207a6e912f.png"
    BackgroundUrl   = "https://storage.googleapis.com/msgsndr/KlBL9XEG0eVNlAqE7m5V/media/6995129b857595d5f3f681bd.jpg"
    
    LocationIcon  = "https://img.icons8.com/material-rounded/24/530607/marker.png"
    PhoneIcon     = "https://img.icons8.com/material-rounded/24/530607/phone.png"
    EmailIcon     = "https://img.icons8.com/material-rounded/24/530607/filled-message.png"
    WebsiteIcon   = "https://img.icons8.com/material-rounded/24/530607/globe.png"
    
    FacebookIcon  = "https://img.icons8.com/material-rounded/24/530607/facebook-new.png"
    InstagramIcon = "https://img.icons8.com/material-rounded/24/530607/instagram-new.png"
    LinkedInIcon  = "https://img.icons8.com/material-rounded/24/530607/linkedin.png"
    WhatsAppIcon  = "https://img.icons8.com/material-rounded/24/530607/whatsapp.png"
    
    FacebookUrl   = "https://www.facebook.com/telexphilippines"
    InstagramUrl  = "https://www.instagram.com/telexph"
    LinkedInUrl   = "https://www.linkedin.com/company/telex-ph"
    WhatsAppUrl   = "https://wa.me/639810996634"
}

function Get-DepartmentSignatureHTML {
    param($DepartmentName, $Email, $Phone, $Address, $BackgroundUrl)

    $googleMapsUrl = "https://www.google.com/maps/place/TELEX+Philippines/@15.6554484,120.7694461,17z/data=!3m1!4b1!4m6!3m5!1s0x33912d233b50b17d:0xff41f0e911207c2!8m2!3d15.6554484!4d120.772021!16s%2Fg%2F11ylzqcd81?entry=ttu&g_ep=EgoyMDI2MDIxMS4wIKXMDSoASAFQAw%3D%3D"
    
    $backgroundStyle = ""
    if ($BackgroundUrl -and $BackgroundUrl -ne "") {
        $backgroundStyle = "background-image:url('$BackgroundUrl');background-size:cover;background-position:center;background-repeat:no-repeat;"
    }
    
    # Department signature: 2 columns only (no photo column)
    # Left: Department name + contact info | Right: Logo + social icons
    $html = "<div style=`"font-family:Arial,sans-serif;color:#000!important;padding:8px 0;width:100%;min-width:480px`">" +
            "<table cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"width:100%;min-width:480px;table-layout:fixed;$backgroundStyle`">" +
            "<tr>" +

            # LEFT COLUMN - Department name + contact info
            "<td style=`"vertical-align:middle;padding:12px 10px;background:rgba(255,255,255,0.82)`">" +
            
            # Department name (big, script font like personal signature)
            "<div style=`"font-family:'Brush Script MT',cursive;font-size:26px;color:#000!important;line-height:1.2;margin-bottom:6px`">$DepartmentName</div>" +
            
            # Red divider line
            "<div style=`"border-bottom:2px solid #530607;margin-bottom:6px;width:100%`"></div>" +
            
            # Contact info
            "<div style=`"border-left:2px solid #000;padding-left:6px`">" +
            "<table cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"font-size:11px;line-height:1.5;width:100%`">" +
            "<tr><td valign=`"top`" style=`"padding-right:4px;padding-bottom:2px;width:20px`"><img src=`"$($config.LocationIcon)`" width=`"14`" height=`"14`" style=`"display:block`" alt=`"Loc`"/></td>" +
            "<td style=`"padding-bottom:2px`"><a href=`"$googleMapsUrl`" target=`"_blank`" style=`"color:#000!important;text-decoration:none`">$Address</a></td></tr>" +
            "<tr><td valign=`"top`" style=`"padding-right:4px;padding-bottom:2px`"><img src=`"$($config.PhoneIcon)`" width=`"14`" height=`"14`" style=`"display:block`" alt=`"Tel`"/></td>" +
            "<td style=`"padding-bottom:2px;color:#000!important`">$Phone</td></tr>" +
            "<tr><td valign=`"top`" style=`"padding-right:4px;padding-bottom:2px`"><img src=`"$($config.EmailIcon)`" width=`"14`" height=`"14`" style=`"display:block`" alt=`"Email`"/></td>" +
            "<td style=`"padding-bottom:2px`"><a href=`"mailto:$Email`" style=`"color:#000!important;text-decoration:none`">$Email</a></td></tr>" +
            "<tr><td valign=`"top`" style=`"padding-right:4px`"><img src=`"$($config.WebsiteIcon)`" width=`"14`" height=`"14`" style=`"display:block`" alt=`"Web`"/></td>" +
            "<td><a href=`"https://$($config.CompanyWebsite)`" style=`"color:#000!important;text-decoration:none`">$($config.CompanyWebsite)</a></td></tr>" +
            "</table></div></td>" +

            # RIGHT COLUMN - Logo + social icons
            "<td style=`"vertical-align:center;text-align:center;padding:20px 6px;width:116px;background:rgba(255,255,255,0.82)`">" +
            "<div style=`"display:inline-block;text-align:center`">" +
            "<img src=`"$($config.LogoUrl)`" width=`"95`" height=`"95`" style=`"display:block;margin:0 auto 3px auto;max-width:100%`" alt=`"TelexPH`"/>" +
            "<table cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"margin:0 auto`"><tr>" +
            "<td style=`"padding:0 2px`"><a href=`"$($config.FacebookUrl)`" style=`"text-decoration:none;display:block`"><img src=`"$($config.FacebookIcon)`" width=`"18`" height=`"18`" style=`"display:block`" alt=`"FB`"/></a></td>" +
            "<td style=`"padding:0 2px`"><a href=`"$($config.InstagramUrl)`" style=`"text-decoration:none;display:block`"><img src=`"$($config.InstagramIcon)`" width=`"18`" height=`"18`" style=`"display:block`" alt=`"IG`"/></a></td>" +
            "<td style=`"padding:0 2px`"><a href=`"$($config.LinkedInUrl)`" style=`"text-decoration:none;display:block`"><img src=`"$($config.LinkedInIcon)`" width=`"18`" height=`"18`" style=`"display:block`" alt=`"LI`"/></a></td>" +
            "<td style=`"padding:0 2px`"><a href=`"$($config.WhatsAppUrl)`" target=`"_blank`" style=`"text-decoration:none;display:block`"><img src=`"$($config.WhatsAppIcon)`" width=`"18`" height=`"18`" style=`"display:block`" alt=`"WA`"/></a></td>" +
            "</tr></table></div></td></tr></table></div>"
    
    return $html
}

function Process-DepartmentUser {
    param([string]$Email)
    Write-Host "`n[PROCESSING DEPARTMENT] $Email" -ForegroundColor Cyan
    Write-Host "=============================================================" -ForegroundColor DarkGray
    
    try {
        $u = Get-User -Identity $Email -ErrorAction Stop
        Write-Host "   [USER] Found: $($u.DisplayName)" -ForegroundColor Green

        # Use DisplayName as department name (set this in M365 admin)
        # e.g. "Human Resources", "Finance Department", "IT Support"
        $departmentName = if ($u.DisplayName) { $u.DisplayName } else { "TelexPH Department" }
        $finalPhone     = if ($u.Phone) { $u.Phone } else { $config.DefaultPhone }
        $finalAddress   = if ($u.StreetAddress) { $u.StreetAddress } else { $config.DefaultAddress }

        Write-Host "   [INFO] Department: $departmentName" -ForegroundColor Gray
        Write-Host "   [INFO] Phone: $finalPhone" -ForegroundColor Gray
        Write-Host "   [INFO] Address: $finalAddress" -ForegroundColor Gray
        
        $finalBackgroundUrl = if ($BackgroundImageUrl -and $BackgroundImageUrl -ne "") { 
            $BackgroundImageUrl 
        } else { 
            $config.BackgroundUrl 
        }

        $html = Get-DepartmentSignatureHTML `
            -DepartmentName $departmentName `
            -Email $Email `
            -Phone $finalPhone `
            -Address $finalAddress `
            -BackgroundUrl $finalBackgroundUrl
            
        Write-Host "   [HTML] Signature length: $($html.Length) characters" -ForegroundColor Cyan
            
        $ruleName = "Email Signature - $($u.DisplayName)"
        
        if (Get-TransportRule $ruleName -ErrorAction SilentlyContinue) {
            Set-TransportRule -Identity $ruleName -ApplyHtmlDisclaimerText $html
            Write-Host "   [RULE] Updated transport rule" -ForegroundColor Green
        } else {
            New-TransportRule -Name $ruleName `
                -From $Email `
                -ApplyHtmlDisclaimerLocation "Append" `
                -ApplyHtmlDisclaimerText $html `
                -ApplyHtmlDisclaimerFallbackAction "Wrap" `
                -Mode "Enforce"
            Write-Host "   [RULE] Created new transport rule" -ForegroundColor Green
        }
        
        Write-Host "   [SUCCESS] Department signature deployed!" -ForegroundColor Magenta
        
    } catch {
        Write-Host "   [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "=============================================================`n" -ForegroundColor DarkGray
}

# --- Main Execution ---
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host "   TelexPH - DEPARTMENT EMAIL SIGNATURE                    " -ForegroundColor Magenta
Write-Host "   - No profile photo                                      " -ForegroundColor Green
Write-Host "   - No job title                                          " -ForegroundColor Green
Write-Host "   - Department name as header                             " -ForegroundColor Green
Write-Host "   - Same background image support                         " -ForegroundColor Green
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host ""

Write-Host "[CLOUDINARY] Cloud Name: $CloudinaryCloudName" -ForegroundColor Gray
Write-Host ""

try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[OK] Connected to Exchange Online" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "[ERROR] NOT CONNECTED TO EXCHANGE ONLINE!" -ForegroundColor Red
    Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
    Write-Host ""
    return
}

if ($UserEmail) {
    Process-DepartmentUser -Email $UserEmail
} elseif ($BatchFile) {
    if (Test-Path $BatchFile) {
        Write-Host "[BATCH] Processing department emails from: $BatchFile`n" -ForegroundColor Cyan
        Get-Content $BatchFile | ForEach-Object { 
            Process-DepartmentUser -Email $_.Trim()
            Start-Sleep -Seconds 2
        }
    } else {
        Write-Host "[ERROR] Batch file not found: $BatchFile" -ForegroundColor Red
    }
} elseif ($AllUsers) {
    Write-Host "[ALL USERS] Processing all mailboxes...`n" -ForegroundColor Cyan
    Get-Mailbox -ResultSize Unlimited | ForEach-Object { 
        Process-DepartmentUser -Email $_.PrimarySmtpAddress
        Start-Sleep -Seconds 2
    }
} else {
    Write-Host "USAGE EXAMPLES:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Single department email:" -ForegroundColor Gray
    Write-Host "  .\ExchangeTransportRuleSignature_DEPARTMENT.ps1 -UserEmail hr@telexph.com -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'key' -CloudinaryApiSecret 'secret'" -ForegroundColor White
    Write-Host ""
    Write-Host "  With custom background:" -ForegroundColor Gray
    Write-Host "  .\ExchangeTransportRuleSignature_DEPARTMENT.ps1 -UserEmail finance@telexph.com -BackgroundImageUrl 'https://...' -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'key' -CloudinaryApiSecret 'secret'" -ForegroundColor White
    Write-Host ""
    Write-Host "  Batch (create a .txt file with one email per line):" -ForegroundColor Gray
    Write-Host "  .\ExchangeTransportRuleSignature_DEPARTMENT.ps1 -BatchFile 'departments.txt' -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'key' -CloudinaryApiSecret 'secret'" -ForegroundColor White
    Write-Host ""
}