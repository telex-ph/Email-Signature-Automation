# TelexPH - Corporate Email Signature (CLOUDINARY VERSION)
# Uploads M365 profile photos to Cloudinary and uses URLs in signatures
# NO BASE64 = NO Exchange size limits = NO ERRORS!
# Faster loading, better quality, CDN delivery

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
    DefaultAddress  = "Cawayan Bugtong, Guimba, Nueva Ecija, Philippines"
    CompanyWebsite  = "www.telexph.com"
    LogoUrl         = "https://storage.googleapis.com/msgsndr/KlBL9XEG0eVNlAqE7m5V/media/69804d9df7a877373924ac5d.png"
    
    # BACKGROUND IMAGE - Red triangular shapes
    BackgroundUrl   = "https://storage.googleapis.com/msgsndr/KlBL9XEG0eVNlAqE7m5V/media/698d3ef952c9527671222933.jpg"
    
    # DEFAULT AVATAR - Gray circle with initials placeholder  
    # Using DARK RED #530607 background
    DefaultPhotoUrl = "https://ui-avatars.com/api/?name=User&size=300&background=530607&color=fff&bold=true"
    
    # ICONS - ALL USING DARK RED #530607 (NOT the old #8B1538!)
    # If icons still show old color, clear Exchange cache or restart Outlook
    LocationIcon  = "https://img.icons8.com/material-rounded/24/530607/marker.png"
    PhoneIcon     = "https://img.icons8.com/material-rounded/24/530607/phone.png"
    EmailIcon     = "https://img.icons8.com/material-rounded/24/530607/filled-message.png"
    WebsiteIcon   = "https://img.icons8.com/material-rounded/24/530607/globe.png"
    
    # Social Media Icons - DARK RED #530607
    FacebookIcon  = "https://img.icons8.com/material-rounded/24/530607/facebook-new.png"
    InstagramIcon = "https://img.icons8.com/material-rounded/24/530607/instagram-new.png"
    LinkedInIcon  = "https://img.icons8.com/material-rounded/24/530607/linkedin.png"
    
    FacebookUrl   = "https://www.facebook.com/telexphilippines"
    InstagramUrl  = "https://www.instagram.com/telexph"
    LinkedInUrl   = "https://www.linkedin.com/company/telex-ph"
}

function Get-SignatureHTML {
    param($DisplayName, $JobTitle, $Email, $Phone, $Address, $PhotoUrl, $BackgroundUrl)
    
    # Build background style if URL provided
    $backgroundStyle = ""
    if ($BackgroundUrl -and $BackgroundUrl -ne "") {
        $backgroundStyle = "background-image:url('$BackgroundUrl');background-size:cover;background-position:center;background-repeat:no-repeat;"
    }
    
    # MULTI-LINE HTML - Easy to read and edit!
    # Settings: 120px circle, 3px border, #530607 color, 20px left padding
    return @"
<div style="font-family:Arial,Helvetica,sans-serif;color:#000;$backgroundStyle">
    <table cellpadding="0" cellspacing="0" border="0" style="background-color:transparent;width:100%;max-width:750px">
        <tr>
            <td style="padding-left:20px;padding-right:25px;vertical-align:middle">
                <div style="width:120px;height:120px;border-radius:50%;border:3px solid #530607;overflow:hidden;display:block;background-color:#f0f0f0">
                    <img src="$PhotoUrl" width="120" height="120" style="object-fit:cover;display:block" alt="Profile Photo"/>
                </div>
            </td>
            <td style="vertical-align:middle;padding-right:30px">
                <div style="font-size:28px;font-weight:900;font-style:italic;color:#000;line-height:1.1;margin-bottom:5px;white-space:nowrap">$DisplayName</div>
                <div style="background-color:#530607;color:#fff;font-size:11px;font-weight:bold;padding:4px 12px;border-radius:4px;display:inline-block;text-transform:uppercase;margin-bottom:15px">$JobTitle</div>
                <div style="border-left:2px solid #000;padding-left:18px">
                    <table cellpadding="0" cellspacing="0" border="0" style="font-size:13px;line-height:1.8">
                        <tr>
                            <td width="22" valign="middle" style="padding-right:8px">
                                <img src="$($config.LocationIcon)" width="16" height="16" style="display:block" alt="Location"/>
                            </td>
                            <td style="padding-bottom:2px">$Address</td>
                        </tr>
                        <tr>
                            <td width="22" valign="middle" style="padding-right:8px">
                                <img src="$($config.PhoneIcon)" width="16" height="16" style="display:block" alt="Phone"/>
                            </td>
                            <td style="padding-bottom:2px">$Phone</td>
                        </tr>
                        <tr>
                            <td width="22" valign="middle" style="padding-right:8px">
                                <img src="$($config.EmailIcon)" width="16" height="16" style="display:block" alt="Email"/>
                            </td>
                            <td style="padding-bottom:2px">
                                <a href="mailto:$Email" style="color:#000;text-decoration:none">$Email</a>
                            </td>
                        </tr>
                        <tr>
                            <td width="22" valign="middle" style="padding-right:8px">
                                <img src="$($config.WebsiteIcon)" width="16" height="16" style="display:block" alt="Website"/>
                            </td>
                            <td>
                                <a href="https://$($config.CompanyWebsite)" style="color:#000;text-decoration:none">$($config.CompanyWebsite)</a>
                            </td>
                        </tr>
                    </table>
                </div>
            </td>
            <td style="vertical-align:middle;text-align:center;padding-left:25px">
                <img src="$($config.LogoUrl)" width="130" style="display:block;margin:0 auto 12px auto" alt="TelexPH Logo"/>
                <div style="white-space:nowrap;text-align:center">
                    <a href="$($config.FacebookUrl)" style="text-decoration:none;display:inline-block;margin:0 4px">
                        <img src="$($config.FacebookIcon)" width="22" height="22" style="display:block" alt="Facebook"/>
                    </a>
                    <a href="$($config.InstagramUrl)" style="text-decoration:none;display:inline-block;margin:0 4px">
                        <img src="$($config.InstagramIcon)" width="22" height="22" style="display:block" alt="Instagram"/>
                    </a>
                    <a href="$($config.LinkedInUrl)" style="text-decoration:none;display:inline-block;margin:0 4px">
                        <img src="$($config.LinkedInIcon)" width="22" height="22" style="display:block" alt="LinkedIn"/>
                    </a>
                </div>
            </td>
        </tr>
    </table>
</div>
"@
}

function Upload-ToCloudinary {
    param([string]$FilePath, [string]$PublicId)
    
    Write-Host "   [CLOUDINARY] Uploading photo..." -ForegroundColor Cyan
    
    # Prepare Cloudinary upload
    $timestamp = [int][double]::Parse((Get-Date -UFormat %s))
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $base64File = [System.Convert]::ToBase64String($fileBytes)
    
    # Create signature
    $stringToSign = "folder=$CloudinaryFolder&public_id=$PublicId&timestamp=$timestamp$CloudinaryApiSecret"
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $hashBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
    $signature = [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()
    
    # Upload to Cloudinary
    $uploadUrl = "https://api.cloudinary.com/v1_1/$CloudinaryCloudName/image/upload"
    
    $body = @{
        file = "data:image/jpeg;base64,$base64File"
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

function Get-UserPhotoUrl {
    param([string]$Email, [string]$DisplayName)
    
    Write-Host "   [PHOTO] Fetching M365 profile photo..." -ForegroundColor Yellow
    
    # Try to get photo from Microsoft Graph
    try {
        # Check if Microsoft.Graph is installed and connected
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) {
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            
            $context = Get-MgContext -ErrorAction SilentlyContinue
            
            if ($context) {
                Write-Host "   [PHOTO] Connected to Microsoft Graph" -ForegroundColor Gray
                
                # Download photo
                $tempFile = [System.IO.Path]::GetTempFileName()
                Get-MgUserPhotoContent -UserId $Email -OutFile $tempFile -ErrorAction Stop
                
                # Generate public_id from email (username part only)
                $publicId = $Email.Split('@')[0]
                
                # Upload to Cloudinary
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
    
    # FALLBACK: Generate avatar with user's initials
    $initials = ($DisplayName -split '\s+' | ForEach-Object { $_[0] }) -join ''
    $encodedName = [System.Uri]::EscapeDataString($DisplayName)
    $fallbackUrl = "https://ui-avatars.com/api/?name=$encodedName&size=300&background=8B1538&color=fff&bold=true"
    
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
        
        # Get photo URL (Cloudinary or fallback)
        $photoUrl = Get-UserPhotoUrl -Email $Email -DisplayName $u.DisplayName
        
        $finalTitle = if ($u.Title) { $u.Title } else { "Team Member" }
        $finalPhone = if ($u.Phone) { $u.Phone } else { $config.DefaultPhone }
        $finalAddress = if ($u.StreetAddress) { $u.StreetAddress } else { $config.DefaultAddress }

        Write-Host "   [INFO] Title: $finalTitle" -ForegroundColor Gray
        Write-Host "   [INFO] Phone: $finalPhone" -ForegroundColor Gray
        Write-Host "   [INFO] Address: $finalAddress" -ForegroundColor Gray
        
        # Use config background URL if parameter not provided
        $finalBackgroundUrl = if ($BackgroundImageUrl -and $BackgroundImageUrl -ne "") { 
            $BackgroundImageUrl 
        } else { 
            $config.BackgroundUrl 
        }
        
        if ($finalBackgroundUrl) {
            Write-Host "   [BACKGROUND] Using background image" -ForegroundColor Cyan
        }
        
        # Display design settings for verification
        Write-Host "   [DESIGN] Circle: 120x120px | Border: 3px | Color: #530607" -ForegroundColor Cyan

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
        
        Write-Host "   [SUCCESS] Email signature deployed!" -ForegroundColor Magenta
        
    } catch {
        Write-Host "   [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "=============================================================`n" -ForegroundColor DarkGray
}

# --- Main Execution ---
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host "   TelexPH Email Signature - WITH BACKGROUND (AUTO)        " -ForegroundColor Magenta
Write-Host "   - Background Image Included (Red Triangles)             " -ForegroundColor Magenta
Write-Host "   - Photos Hosted on Cloudinary CDN                       " -ForegroundColor Magenta
Write-Host "   - NO Base64 = NO Size Limits                            " -ForegroundColor Magenta
Write-Host "   - Fast Loading, Professional Design                     " -ForegroundColor Magenta
Write-Host "   - GUARANTEED No Exchange Errors!                        " -ForegroundColor Magenta
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host ""

# Validate Cloudinary credentials
Write-Host "[CLOUDINARY] Validating credentials..." -ForegroundColor Cyan
Write-Host "   Cloud Name: $CloudinaryCloudName" -ForegroundColor Gray
Write-Host "   API Key: $CloudinaryApiKey" -ForegroundColor Gray
Write-Host ""

# Check Exchange Connection
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

# Check Graph Connection
$context = Get-MgContext -ErrorAction SilentlyContinue
if ($context) {
    Write-Host "[OK] Connected to Microsoft Graph" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "[WARNING] Not connected to Microsoft Graph - will use fallback avatars" -ForegroundColor Yellow
    Write-Host ""
}

# Run Logic
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
    Write-Host "  Single user (uses background from config):" -ForegroundColor White
    Write-Host "    .\CloudinarySignature_WithBackground.ps1 -UserEmail hjreyes@telexph.com -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'your-key' -CloudinaryApiSecret 'your-secret'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Batch file (all users get background):" -ForegroundColor White
    Write-Host "    .\CloudinarySignature_WithBackground.ps1 -BatchFile users.txt -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'your-key' -CloudinaryApiSecret 'your-secret'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  All users:" -ForegroundColor White
    Write-Host "    .\CloudinarySignature_WithBackground.ps1 -AllUsers -CloudinaryCloudName 'dzwxer9cq' -CloudinaryApiKey 'your-key' -CloudinaryApiSecret 'your-secret'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Optional: Override background URL:" -ForegroundColor White
    Write-Host "    Add: -BackgroundImageUrl 'https://different-background.jpg'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "NOTE: Background URL is set in config. Edit script to change default background." -ForegroundColor Yellow
    Write-Host ""
}