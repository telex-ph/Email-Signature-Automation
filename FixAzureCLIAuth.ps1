# TelexPH - AUTOMATED M365 Photo Sync (NO MANUAL UPLOAD!)
# Fixes Azure CLI authentication and enables automatic photo fetching

Write-Host ""
Write-Host "================================================================" -ForegroundColor Magenta
Write-Host "   AUTOMATED M365 PHOTO SYNC - FIX AUTHENTICATION             " -ForegroundColor Magenta
Write-Host "   Automatic fetch from M365 - NO manual upload!              " -ForegroundColor Magenta
Write-Host "================================================================" -ForegroundColor Magenta
Write-Host ""

# Check if Azure CLI is installed
Write-Host "[CHECK] Checking if Azure CLI is installed..." -ForegroundColor Yellow

try {
    $azVersion = az --version 2>$null
    if ($azVersion) {
        Write-Host "[OK] Azure CLI is installed" -ForegroundColor Green
    } else {
        Write-Host "[ERROR] Azure CLI not found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "INSTALL FIRST:" -ForegroundColor Yellow
        Write-Host "  Download: https://aka.ms/installazurecliwindows" -ForegroundColor White
        Write-Host "  Install MSI, then restart PowerShell" -ForegroundColor White
        Write-Host ""
        exit
    }
} catch {
    Write-Host "[ERROR] Azure CLI not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "INSTALL FIRST:" -ForegroundColor Yellow
    Write-Host "  Download: https://aka.ms/installazurecliwindows" -ForegroundColor White
    Write-Host "  Install MSI, then restart PowerShell" -ForegroundColor White
    Write-Host ""
    exit
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "   3 AUTHENTICATION METHODS (Choose one that works)           " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Method 1: Browser Login (default)" -ForegroundColor Yellow
Write-Host "  - Opens browser for authentication" -ForegroundColor Gray
Write-Host "  - Most common method" -ForegroundColor Gray
Write-Host ""

Write-Host "Method 2: Device Code Flow" -ForegroundColor Yellow
Write-Host "  - Use different device/browser" -ForegroundColor Gray
Write-Host "  - Good for terminal issues" -ForegroundColor Gray
Write-Host ""

Write-Host "Method 3: Service Principal (advanced)" -ForegroundColor Yellow
Write-Host "  - For automation/scripts" -ForegroundColor Gray
Write-Host "  - Requires app registration" -ForegroundColor Gray
Write-Host ""

$method = Read-Host "Choose method (1/2/3) or press Enter for Method 1"

if ([string]::IsNullOrWhiteSpace($method)) {
    $method = "1"
}

Write-Host ""

switch ($method) {
    "1" {
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "   METHOD 1: Browser Login                                    " -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host ""
        
        Write-Host "[INFO] This will open a browser window for authentication" -ForegroundColor Gray
        Write-Host "[INFO] If browser doesn't open or hangs, use Method 2 instead" -ForegroundColor Gray
        Write-Host ""
        
        # Clear any existing login
        Write-Host "[CLEANUP] Clearing existing Azure CLI session..." -ForegroundColor Cyan
        az logout 2>$null
        
        Write-Host "[LOGIN] Opening browser for authentication..." -ForegroundColor Yellow
        Write-Host ""
        
        try {
            # Try standard browser login
            az login --use-device-code:$false
            
            # Verify login
            $account = az account show 2>$null | ConvertFrom-Json
            
            if ($account) {
                Write-Host ""
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host "   ✅ LOGIN SUCCESS!                                          " -ForegroundColor Green
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host ""
                Write-Host "[INFO] Logged in as: $($account.user.name)" -ForegroundColor Gray
                Write-Host "[INFO] Tenant: $($account.tenantId)" -ForegroundColor Gray
                Write-Host ""
            } else {
                throw "Login verification failed"
            }
            
        } catch {
            Write-Host ""
            Write-Host "[ERROR] Browser login failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "TRY METHOD 2 INSTEAD:" -ForegroundColor Yellow
            Write-Host "  Run this script again and choose option 2" -ForegroundColor White
            Write-Host ""
            exit
        }
    }
    
    "2" {
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "   METHOD 2: Device Code Flow                                 " -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host ""
        
        Write-Host "[INFO] This method works even if browser has issues" -ForegroundColor Gray
        Write-Host "[INFO] You'll get a code to enter at a URL" -ForegroundColor Gray
        Write-Host ""
        
        # Clear any existing login
        Write-Host "[CLEANUP] Clearing existing Azure CLI session..." -ForegroundColor Cyan
        az logout 2>$null
        
        Write-Host "[LOGIN] Starting device code authentication..." -ForegroundColor Yellow
        Write-Host ""
        
        try {
            # Use device code flow
            az login --use-device-code
            
            # Verify login
            $account = az account show 2>$null | ConvertFrom-Json
            
            if ($account) {
                Write-Host ""
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host "   ✅ LOGIN SUCCESS!                                          " -ForegroundColor Green
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host ""
                Write-Host "[INFO] Logged in as: $($account.user.name)" -ForegroundColor Gray
                Write-Host "[INFO] Tenant: $($account.tenantId)" -ForegroundColor Gray
                Write-Host ""
            } else {
                throw "Login verification failed"
            }
            
        } catch {
            Write-Host ""
            Write-Host "[ERROR] Device code login failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            exit
        }
    }
    
    "3" {
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "   METHOD 3: Service Principal (Advanced)                     " -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host ""
        
        Write-Host "[INFO] This requires Azure AD App Registration" -ForegroundColor Gray
        Write-Host "[INFO] For automated scripts without user interaction" -ForegroundColor Gray
        Write-Host ""
        
        $appId = Read-Host "Enter Application (Client) ID"
        $tenantId = Read-Host "Enter Tenant ID"
        $secret = Read-Host "Enter Client Secret" -AsSecureString
        
        $secretPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
        )
        
        Write-Host ""
        Write-Host "[LOGIN] Authenticating with service principal..." -ForegroundColor Yellow
        
        try {
            az login --service-principal -u $appId -p $secretPlainText --tenant $tenantId
            
            # Verify login
            $account = az account show 2>$null | ConvertFrom-Json
            
            if ($account) {
                Write-Host ""
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host "   ✅ LOGIN SUCCESS!                                          " -ForegroundColor Green
                Write-Host "================================================================" -ForegroundColor Green
                Write-Host ""
                Write-Host "[INFO] Logged in as: $appId" -ForegroundColor Gray
                Write-Host "[INFO] Tenant: $tenantId" -ForegroundColor Gray
                Write-Host ""
            } else {
                throw "Login verification failed"
            }
            
        } catch {
            Write-Host ""
            Write-Host "[ERROR] Service principal login failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            exit
        }
    }
    
    default {
        Write-Host "[ERROR] Invalid choice!" -ForegroundColor Red
        exit
    }
}

# Now test if we can access Microsoft Graph
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "   Testing Microsoft Graph API Access                         " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[TEST] Getting access token for Microsoft Graph..." -ForegroundColor Yellow

try {
    $tokenResult = az account get-access-token --resource https://graph.microsoft.com 2>$null | ConvertFrom-Json
    
    if ($tokenResult -and $tokenResult.accessToken) {
        Write-Host "[OK] Graph API access token obtained!" -ForegroundColor Green
        Write-Host "[INFO] Token expires: $($tokenResult.expiresOn)" -ForegroundColor Gray
        Write-Host ""
    } else {
        throw "Failed to get access token"
    }
} catch {
    Write-Host "[ERROR] Failed to get Graph API token: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "You may need to:" -ForegroundColor Yellow
    Write-Host "  1. Grant Microsoft Graph permissions" -ForegroundColor White
    Write-Host "  2. Contact your Azure AD admin" -ForegroundColor White
    Write-Host ""
    exit
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "   ✅ SETUP COMPLETE! READY FOR AUTOMATED SYNC                " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "NOW YOU CAN:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Sync M365 photos AUTOMATICALLY:" -ForegroundColor White
Write-Host "   .\ForceSyncM365Photo.ps1 -UserEmail hjreyes@telexph.com" -ForegroundColor Cyan
Write-Host ""
Write-Host "2. Deploy email signatures:" -ForegroundColor White
Write-Host "   .\ExchangeTransportRuleSignature.ps1 -UserEmail hjreyes@telexph.com" -ForegroundColor Cyan
Write-Host ""
Write-Host "3. For ALL users (batch processing):" -ForegroundColor White
Write-Host "   Get-Mailbox | ForEach-Object {" -ForegroundColor Cyan
Write-Host "       .\ForceSyncM365Photo.ps1 -UserEmail $_.PrimarySmtpAddress" -ForegroundColor Cyan
Write-Host "       .\ExchangeTransportRuleSignature.ps1 -UserEmail $_.PrimarySmtpAddress" -ForegroundColor Cyan
Write-Host "   }" -ForegroundColor Cyan
Write-Host ""
Write-Host "NO MORE MANUAL UPLOADS! ALL AUTOMATED! ✅" -ForegroundColor Green
Write-Host ""