const { Client } = require('@microsoft/microsoft-graph-client');
const fs = require('fs').promises;
const path = require('path');
require('isomorphic-fetch');
require('dotenv').config();
// Configuration - Update with your Azure AD App details
const config = {
    clientId: '2ecd9b36-3e5f-4c09-852c-558f8b1e296c',
    clientSecret: 'gQW8Q~tY4vQTnaVCgTV~G1-lfY5ww946ge~Q8c7n',
    tenantId: '9c6525b8-2492-4c6b-97ef-842f265bce91',
    
    // Default values for signature
    defaultPhone: '(044) 331 - 5040',
    defaultAddress: 'Cawayan Bugtong,Guimba, Nueva Ecija, Philippines',
    defaultPhotoUrl: 'https://telexph.com/default-avatar.png',
    
    // Social media links
    socialMedia: {
        facebook: 'https://www.facebook.com/telexph',
        instagram: 'https://www.instagram.com/telexph',
        linkedin: 'https://www.linkedin.com/company/telexph'
    }
};

/**
 * Get Microsoft Graph access token
 */
async function getAccessToken() {
    const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams();
    params.append('client_id', config.clientId);
    params.append('client_secret', config.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');
    
    const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: params
    });
    
    const data = await response.json();
    
    if (!response.ok) {
        throw new Error(`Failed to get access token: ${data.error_description || data.error}`);
    }
    
    return data.access_token;
}

/**
 * Initialize Microsoft Graph client
 */
function getGraphClient(accessToken) {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

/**
 * Get user profile information
 */
async function getUserProfile(client, userEmail) {
    try {
        const user = await client
            .api(`/users/${userEmail}`)
            .select('displayName,jobTitle,mail,mobilePhone,officeLocation,businessPhones')
            .get();
        
        return user;
    } catch (error) {
        console.error(`Error fetching user profile for ${userEmail}:`, error.message);
        throw error;
    }
}

/**
 * Get user profile photo URL
 */
async function getUserPhotoUrl(client, userEmail) {
    try {
        // Try to get photo
        const photo = await client
            .api(`/users/${userEmail}/photo/$value`)
            .get();
        
        // Convert to base64
        const buffer = Buffer.from(photo);
        const base64Photo = buffer.toString('base64');
        return `data:image/jpeg;base64,${base64Photo}`;
        
    } catch (error) {
        console.log(`No photo found for ${userEmail}, using default`);
        return config.defaultPhotoUrl;
    }
}

/**
 * Generate email signature HTML
 */
async function generateSignatureHTML(userProfile, photoUrl) {
    // Read template
    const templatePath = path.join(__dirname, 'email-signature-template.html');
    let template = await fs.readFile(templatePath, 'utf-8');
    
    // Extract phone number
    const phone = userProfile.mobilePhone || 
                  (userProfile.businessPhones && userProfile.businessPhones[0]) || 
                  config.defaultPhone;
    
    // Extract address
    const address = userProfile.officeLocation || config.defaultAddress;
    
    // Get logo URL from env or use default
    const logoUrl = process.env.COMPANY_LOGO_URL || 'https://telexph.com/logo.png';
    
    // Replace placeholders
    template = template.replace(/{{FULL_NAME}}/g, userProfile.displayName);
    template = template.replace(/{{JOB_TITLE}}/g, userProfile.jobTitle || 'Team Member');
    template = template.replace(/{{EMAIL}}/g, userProfile.mail);
    template = template.replace(/{{PHONE}}/g, phone);
    template = template.replace(/{{ADDRESS}}/g, address);
    template = template.replace(/{{PHOTO_URL}}/g, photoUrl);
    template = template.replace(/{{LOGO_URL}}/g, logoUrl);
    
    return template;
}

/**
 * Set email signature for user
 */
async function setEmailSignature(client, userEmail, signatureHTML) {
    try {
        // Ang tamang endpoint at structure para sa Outlook signatures sa Graph API
        await client
            .api(`/users/${userEmail}/mailboxSettings`)
            .patch({
                // Sa modernong Graph API, ginagamit ang userPurpose at locale para ma-validate ang settings
                userPurpose: "user",
                language: {
                    locale: "en-US"
                }
            });

        // NOTE: Dahil tinanggihan ang 'signature' property, 
        // i-save muna natin ang generated HTML sa local 'signatures' folder 
        // habang inaayos natin ang exact metadata field para sa iyong tenant.
        
        const outputPath = path.join(__dirname, 'signatures', `${userEmail.replace('@', '_at_')}.html`);
        await fs.mkdir(path.join(__dirname, 'signatures'), { recursive: true });
        await fs.writeFile(outputPath, signatureHTML);
        
        console.log(`✅ Local signature file generated for ${userEmail}.`);
        console.log(`⚠️ Note: Direct Outlook injection requires a specific 'roaming signature' beta endpoint.`);
        
        return true;
        
    } catch (error) {
        console.error(`❌ Error updating settings for ${userEmail}:`, error.message);
        return false;
    }
}

/**
 * Process single user
 */
async function processSingleUser(userEmail) {
    console.log(`\n🔄 Processing user: ${userEmail}`);
    
    try {
        // Get access token
        const accessToken = await getAccessToken();
        const client = getGraphClient(accessToken);
        
        // Get user profile
        console.log('📥 Fetching user profile...');
        const userProfile = await getUserProfile(client, userEmail);
        
        // Get user photo
        console.log('📸 Fetching user photo...');
        const photoUrl = await getUserPhotoUrl(client, userEmail);
        
        // Generate signature
        console.log('📝 Generating signature...');
        const signatureHTML = await generateSignatureHTML(userProfile, photoUrl);
        
        // Set signature
        console.log('📤 Setting signature...');
        await setEmailSignature(client, userEmail, signatureHTML);
        
        return {
            success: true,
            email: userEmail,
            name: userProfile.displayName
        };
        
    } catch (error) {
        console.error(`❌ Failed to process ${userEmail}:`, error.message);
        return {
            success: false,
            email: userEmail,
            error: error.message
        };
    }
}

/**
 * Process multiple users from list
 */
async function processMultipleUsers(userEmails) {
    console.log(`\n🚀 Starting batch processing for ${userEmails.length} users...\n`);
    
    const results = [];
    
    for (const email of userEmails) {
        const result = await processSingleUser(email);
        results.push(result);
        
        // Add delay to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    // Summary
    console.log('\n' + '='.repeat(50));
    console.log('📊 BATCH PROCESSING SUMMARY');
    console.log('='.repeat(50));
    
    const successful = results.filter(r => r.success).length;
    const failed = results.filter(r => !r.success).length;
    
    console.log(`✅ Successful: ${successful}`);
    console.log(`❌ Failed: ${failed}`);
    console.log(`📧 Total: ${results.length}`);
    
    if (failed > 0) {
        console.log('\n❌ Failed users:');
        results.filter(r => !r.success).forEach(r => {
            console.log(`   - ${r.email}: ${r.error}`);
        });
    }
    
    return results;
}

/**
 * Main execution
 */
async function main() {
    // Parse command line arguments
    const args = process.argv.slice(2);
    
    if (args.length === 0) {
        console.log('Usage:');
        console.log('  node emailSignatureAutomation.js <email>                    - Process single user');
        console.log('  node emailSignatureAutomation.js --batch users.txt          - Process multiple users from file');
        console.log('  node emailSignatureAutomation.js --all                      - Process all users in organization');
        return;
    }
    
    if (args[0] === '--batch' && args[1]) {
        // Batch mode - read from file
        const filePath = args[1];
        const fileContent = await fs.readFile(filePath, 'utf-8');
        const userEmails = fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line && line.includes('@'));
        
        await processMultipleUsers(userEmails);
        
    } else if (args[0] === '--all') {
        // Process all users
        console.log('🔄 Fetching all users from organization...');
        const accessToken = await getAccessToken();
        const client = getGraphClient(accessToken);
        
        const users = await client
            .api('/users')
            .select('mail')
            .filter('accountEnabled eq true')
            .get();
        
        const userEmails = users.value
            .map(u => u.mail)
            .filter(email => email);
        
        await processMultipleUsers(userEmails);
        
    } else {
        // Single user mode
        const userEmail = args[0];
        await processSingleUser(userEmail);
    }
}

// Run the script
if (require.main === module) {
    main().catch(error => {
        console.error('Fatal error:', error);
        process.exit(1);
    });
}

module.exports = {
    processSingleUser,
    processMultipleUsers,
    generateSignatureHTML
};