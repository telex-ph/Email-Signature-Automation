# TelexPH - Corporate Email Signature (MOBILE-FRIENDLY)
# - CLOSER SPACING between photo and text column!
# - Photo: 95px, Logo: 105px (optimized for mobile!)
# - Name Font: 20px (NO WRAP on mobile!)
# - REDUCED padding: 10px 3px (instead of 10px 10px)
# - Contact Icons: 16px, Social Icons: 19px
# - Text: 11px (readable on small screens)
# - Column widths: 22% | 52% | 26%
# - Table: table-layout:fixed, max-width 720px (NO min-width = responsive!)
# - Wave: Unicode WAVE DASH chars (〜) in #530607 - works in all email clients

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
    DefaultPhotoUrl = "https://ui-avatars.com/api/?name=User&size=300&background=530607&color=fff&bold=true"
    
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

function Get-SignatureHTML {
    param($DisplayName, $JobTitle, $Email, $Phone, $Address, $PhotoUrl, $BackgroundUrl)

    $googleMapsUrl = "https://www.google.com/maps/place/TELEX+Philippines/@15.6554484,120.7694461,17z/data=!3m1!4b1!4m6!3m5!1s0x33912d233b50b17d:0xff41f0e911207c2!8m2!3d15.6554484!4d120.772021!16s%2Fg%2F11ylzqcd81?entry=ttu&g_ep=EgoyMDI2MDIxMS4wIKXMDSoASAFQAw%3D%3D"
    
    $backgroundStyle = ""
    if ($BackgroundUrl -and $BackgroundUrl -ne "") {
        $backgroundStyle = "background-image:url('$BackgroundUrl');background-size:cover;background-position:center;background-repeat:no-repeat;"
    }
    
    $html = "<div style=`"font-family:Arial,sans-serif;color:#000!important;padding:8px 0;width:100%;min-width:480px`">" +
            "<table cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"width:100%;min-width:480px;table-layout:fixed;$backgroundStyle`">" +
            "<tr>" +
            "<td style=`"padding:10px 6px;vertical-align:middle;width:106px;background:rgba(255,255,255,0.82)`">" +
            "<div style=`"width:90px;height:90px;border-radius:50%;border:2px solid #530607;overflow:hidden;display:block;background:#f0f0f0;margin:0 auto`">" +
            "<img src=`"$PhotoUrl`" width=`"90`" height=`"90`" style=`"object-fit:cover;display:block;width:100%;height:100%`" alt=`"Photo`"/></div></td>" +
            "<td style=`"vertical-align:middle;padding:10px 4px;background:rgba(255,255,255,0.82)`">"+
            "<div style=`"font-family:'Brush Script MT',cursive;font-size:28px;color:#000!important;line-height:1.2;margin-bottom:4px`">$DisplayName</div>" +
            "<div style=`"background:#530607;color:#fff;font-size:9px;font-weight:bold;padding:2px 7px;border-radius:3px;display:inline-block;text-transform:uppercase;margin-bottom:5px`">$JobTitle</div>" +
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

function Upload-ToCloudinary {
    param([string]$FilePath, [string]$PublicId)
    
    Write-Host "   [CLOUDINARY] Uploading photo..." -ForegroundColor Cyan
    
    $timestamp = [int][double]::Parse((Get-Date -UFormat %s))
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $base64File = [System.Convert]::ToBase64String($fileBytes)
    
    $stringToSign = "folder=$CloudinaryFolder&public_id=$PublicId&timestamp=$timestamp$CloudinaryApiSecret"
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $hashBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
    $signature = [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()
    
    $uploadUrl = "https://api.cloudinary.com/v1_1/$CloudinaryCloudName/image/upload"
    
    # FIX: Use format operator to avoid PowerShell comma parsing issue
    $fileDataUrl = ("data:image/jpeg;base64,{0}" -f $base64File)
    
    $body = @{
        file = $fileDataUrl
        public_id = $PublicId
        folder = $CloudinaryFolder
        timestamp = $timestamp
        api_key = $CloudinaryApiKey
        signature = $signature
    }
    
    try {
        $response = Invoke-RestMethod -Uri $uploadUrl -Method Post -Body $body
        $photoUrl = $response.secure_url
        Write-Host "   [CLOUDINARY] SUCCESS! Photo uploaded" -ForegroundColor Green
        Write-Host "   [CLOUDINARY] URL: $photoUrl" -ForegroundColor Gray
        return $photoUrl
    } catch {
        Write-Host "   [CLOUDINARY] Upload failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Test-CloudinaryUrlExists {
    param([string]$Url)
    try {
        $response = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Get-UserPhotoUrl {
    param([string]$Email, [string]$DisplayName)
    
    Write-Host "   [PHOTO] Fetching M365 profile photo..." -ForegroundColor Yellow
    
    $publicId = $Email.Split('@')[0]
    
    # Step 1: Check if photo already exists in Cloudinary - skip re-upload if it does
    $existingUrl = "https://res.cloudinary.com/$CloudinaryCloudName/image/upload/email-signatures/$publicId.jpg"
    Write-Host "   [PHOTO] Checking existing Cloudinary URL..." -ForegroundColor Gray
    if (Test-CloudinaryUrlExists -Url $existingUrl) {
        Write-Host "   [PHOTO] Found existing Cloudinary photo - reusing it!" -ForegroundColor Green
        Write-Host "   [PHOTO] URL: $existingUrl" -ForegroundColor Gray
        return $existingUrl
    }
    Write-Host "   [PHOTO] No existing photo found - will upload from M365" -ForegroundColor Yellow

    # Step 2: Try to get from M365 and upload to Cloudinary
    try {
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) {
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            
            $context = Get-MgContext -ErrorAction SilentlyContinue
            
            if ($context) {
                Write-Host "   [PHOTO] Connected to Microsoft Graph" -ForegroundColor Gray
                
                $tempFile = [System.IO.Path]::GetTempFileName()
                Get-MgUserPhotoContent -UserId $Email -OutFile $tempFile -ErrorAction Stop
                
                $photoUrl = Upload-ToCloudinary -FilePath $tempFile -PublicId $publicId
                
                Remove-Item $tempFile -Force
                
                if ($photoUrl) {
                    return $photoUrl
                }
            } else {
                Write-Host "   [PHOTO] Not connected to Microsoft Graph" -ForegroundColor DarkGray
            }
        } else {
            Write-Host "   [PHOTO] Microsoft.Graph.Users module not installed" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "   [PHOTO] Error: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
    
    $initials = ($DisplayName -split '\s+' | ForEach-Object { $_[0] }) -join ''
    $encodedName = [System.Uri]::EscapeDataString($DisplayName)
    $fallbackUrl = "https://ui-avatars.com/api/?name=$encodedName&size=300&background=530607&color=fff&bold=true"
    
    Write-Host "   [PHOTO] Using fallback avatar with initials: $initials" -ForegroundColor Cyan
    return $fallbackUrl
}

function Process-SingleUser {
    param([string]$Email)
    Write-Host "`n[PROCESSING] $Email" -ForegroundColor Cyan
    Write-Host "=============================================================" -ForegroundColor DarkGray
    
    try {
        $u = Get-User -Identity $Email -ErrorAction Stop
        Write-Host "   [USER] Found: $($u.DisplayName)" -ForegroundColor Green
        
        $photoUrl = Get-UserPhotoUrl -Email $Email -DisplayName $u.DisplayName
        
        $finalTitle = if ($u.Title) { $u.Title } else { "Team Member" }
        $finalPhone = if ($u.Phone) { $u.Phone } else { $config.DefaultPhone }
        $finalAddress = if ($u.StreetAddress) { $u.StreetAddress } else { $config.DefaultAddress }

        Write-Host "   [INFO] Title: $finalTitle" -ForegroundColor Gray
        Write-Host "   [INFO] Phone: $finalPhone" -ForegroundColor Gray
        Write-Host "   [INFO] Address: $finalAddress" -ForegroundColor Gray
        
        $finalBackgroundUrl = if ($BackgroundImageUrl -and $BackgroundImageUrl -ne "") { 
            $BackgroundImageUrl 
        } else { 
            $config.BackgroundUrl 
        }
        
        if ($finalBackgroundUrl) {
            Write-Host "   [BACKGROUND] Using background image" -ForegroundColor Cyan
        }
        
        Write-Host "   [DESIGN] Photo: 95px | Logo: 105px | Border: 3px #530607" -ForegroundColor Cyan
        Write-Host "   [SPACING] TEXT CLOSER TO PHOTO! (3px vs 10px before)" -ForegroundColor Green
        Write-Host "   [FONT] Name: 20px (NO WRAP on mobile!), Text: 11px" -ForegroundColor Magenta
        Write-Host "   [ICONS] Contact: 16px, Social: 19px - Mobile-optimized!" -ForegroundColor Cyan

        $html = Get-SignatureHTML `
            -DisplayName $u.DisplayName `
            -JobTitle $finalTitle `
            -Email $Email `
            -Phone $finalPhone `
            -Address $finalAddress `
            -PhotoUrl $photoUrl `
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
        
        Write-Host "   [SUCCESS] Closer spacing signature deployed!" -ForegroundColor Magenta
        
    } catch {
        Write-Host "   [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "=============================================================`n" -ForegroundColor DarkGray
}

# --- Main Execution ---
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host "   TelexPH - CLOSER SPACING VERSION!                       " -ForegroundColor Magenta
Write-Host "   - Text CLOSER to photo (3px left padding!)              " -ForegroundColor Green
Write-Host "   - Was: 10px | Now: 3px = 70% CLOSER!                    " -ForegroundColor Green
Write-Host "   - Tighter, more compact design                          " -ForegroundColor Green
Write-Host "   - PowerShell comma error FIXED!                         " -ForegroundColor Green
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host ""

Write-Host "[CLOUDINARY] Validating credentials..." -ForegroundColor Cyan
Write-Host "   Cloud Name: $CloudinaryCloudName" -ForegroundColor Gray
Write-Host "   API Key: $CloudinaryApiKey" -ForegroundColor Gray
Write-Host ""

try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[OK] Connected to Exchange Online" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "[ERROR] NOT CONNECTED TO EXCHANGE ONLINE!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please run: Connect-ExchangeOnline" -ForegroundColor Yellow
    Write-Host ""
    return
}

$context = Get-MgContext -ErrorAction SilentlyContinue
if ($context) {
    Write-Host "[OK] Connected to Microsoft Graph" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "[WARNING] Not connected to Microsoft Graph - will use fallback avatars" -ForegroundColor Yellow
    Write-Host ""
}

if ($UserEmail) {
    Process-SingleUser -Email $UserEmail
} elseif ($BatchFile) {
    if (Test-Path $BatchFile) {
        Write-Host "[BATCH] Processing users from file: $BatchFile`n" -ForegroundColor Cyan
        Get-Content $BatchFile | ForEach-Object { 
            Process-SingleUser -Email $_.Trim()
            Start-Sleep -Seconds 2
        }
    } else {
        Write-Host "[ERROR] Batch file not found: $BatchFile" -ForegroundColor Red
    }
} elseif ($AllUsers) {
    Write-Host "[ALL USERS] Processing all mailboxes...`n" -ForegroundColor Cyan
    Get-Mailbox -ResultSize Unlimited | ForEach-Object { 
        Process-SingleUser -Email $_.PrimarySmtpAddress
        Start-Sleep -Seconds 2
    }
} else {
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\ExchangeTransportRuleSignature_CLOSER.ps1 -UserEmail hjreyes@telexph.com -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'key' -CloudinaryApiSecret 'secret'" -ForegroundColor Gray
    Write-Host ""
}