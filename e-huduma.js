// Configuration
var CONFIG = {
  SHEET_NAME: 'E-Huduma_Users',
  SERVICE_NAME: 'E-Huduma',
  USERNAME_PREFIX: 'EH'
};

// Initialize the system
function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // Create headers
    var headers = [
      'User Number', 'Username', 'Email', 'Password', 'Name', 
      'Location Address', 'Phone Number', 'Date of Birth', 
      'Gender', 'Nationality', 'Date Registered', 'Last Updated'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Generate UNIQUE username - FIXED VERSION
function generateUsername(dateRegistered) {
  try {
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    // Find the highest user number
    var maxUserNumber = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && !isNaN(data[i][0])) {
        maxUserNumber = Math.max(maxUserNumber, parseInt(data[i][0]));
      }
    }
    
    var nextUserNumber = maxUserNumber + 1;
    var dateStr = Utilities.formatDate(dateRegistered, Session.getScriptTimeZone(), 'yyyyMMdd');
    var userNumStr = nextUserNumber.toString().padStart(3, '0');
    
    return CONFIG.USERNAME_PREFIX + '-' + dateStr + '-' + userNumStr;
  } catch (error) {
    Logger.log('Username generation error: ' + error.toString());
    // Fallback if there's an error
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    return CONFIG.USERNAME_PREFIX + '-' + dateStr + '-001';
  }
}

// Generate random password
function generatePassword() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%';
  var password = '';
  for (var i = 0; i < 10; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

// Register new user - FIXED with proper user number assignment
function registerUser(userData) {
  try {
    var sheet = initializeSystem();
    var lastRow = sheet.getLastRow();
    
    // Calculate next user number correctly
    var userNumber = lastRow === 0 ? 1 : lastRow; // Header row doesn't count as user
    
    // Generate credentials
    var dateRegistered = new Date();
    var password = generatePassword();
    var username = generateUsername(dateRegistered);
    
    // Prepare user record
    var userRecord = [
      userNumber,
      username,
      userData.email,
      password,
      userData.name,
      userData.address,
      userData.phone,
      userData.dob,
      userData.gender,
      userData.nationality,
      dateRegistered,
      dateRegistered
    ];
    
    // Save to sheet
    sheet.getRange(lastRow + 1, 1, 1, userRecord.length).setValues([userRecord]);
    
    // Send confirmation email
    sendWelcomeEmail(userData.email, username, password, userData.name);
    
    return {
      success: true,
      username: username,
      message: 'Registration successful! Check your email for credentials.'
    };
    
  } catch (error) {
    Logger.log('Registration error: ' + error.toString());
    return {
      success: false,
      message: 'Error during registration: ' + error.toString()
    };
  }
}

// Send welcome email
function sendWelcomeEmail(email, username, password, name) {
  try {
    var subject = 'Welcome to ' + CONFIG.SERVICE_NAME + ' - Your Account Details';
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #072033;">Welcome to ${CONFIG.SERVICE_NAME}, ${name}!</h2>
        <p>Your account has been successfully created. Here are your login details:</p>
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
          <p><strong>Username:</strong> ${username}</p>
          <p><strong>Password:</strong> ${password}</p>
        </div>
        <p style="color: #666; font-size: 14px;">
          <strong>Important:</strong> Please keep your credentials secure and change your password after first login.
        </p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 30px 0;">
        <p style="color: #999; font-size: 12px;">
          This is an automated message from E-Huduma. Please do not reply to this email.
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (error) {
    Logger.log('Email error: ' + error.toString());
  }
}

// Get user data
function getUserData(username) {
  try {
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    // Skip header row (index 0)
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === username.trim()) {
        Logger.log('User found: ' + username);
        return {
          success: true,
          data: {
            username: data[i][1],
            email: data[i][2],
            name: data[i][4],
            address: data[i][5],
            phone: data[i][6],
            dob: data[i][7] ? formatDateForDisplay(data[i][7]) : '',
            gender: data[i][8],
            nationality: data[i][9],
            dateRegistered: data[i][10] ? formatDateForDisplay(data[i][10]) : ''
          }
        };
      }
    }
    Logger.log('User not found: ' + username);
    return { success: false, message: 'User not found. Please check your username.' };
  } catch (error) {
    Logger.log('Get user error: ' + error.toString());
    return { success: false, message: 'Error retrieving user data: ' + error.toString() };
  }
}

// Change password function - SIMPLIFIED VERSION
function changePassword(username, currentPassword, newPassword) {
  try {
    Logger.log('Change password called for: ' + username);
    
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === username.trim()) {
        // Check if current password matches
        if (data[i][3] !== currentPassword) {
          return { success: false, message: 'Current password is incorrect' };
        }
        
        var row = i + 1;
        
        // Update password
        sheet.getRange(row, 4).setValue(newPassword);
        sheet.getRange(row, 12).setValue(new Date()); // Update timestamp
        
        // Send password change confirmation email
        sendPasswordChangeEmail(data[i][2], username, data[i][4]);
        
        Logger.log('Password changed successfully for: ' + username);
        return { success: true, message: 'Password changed successfully!' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (error) {
    Logger.log('Password change error: ' + error.toString());
    return { success: false, message: 'Error changing password: ' + error.toString() };
  }
}

// Reset password - SIMPLIFIED VERSION
function resetPassword(username, email) {
  try {
    Logger.log('Reset password called for: ' + username);
    
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === username.trim() && 
          data[i][2].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
        
        var newPassword = generatePassword();
        var row = i + 1;
        
        // Update password
        sheet.getRange(row, 4).setValue(newPassword);
        sheet.getRange(row, 12).setValue(new Date());
        
        // Send new password email
        sendPasswordResetEmail(data[i][2], username, newPassword, data[i][4]);
        
        Logger.log('Password reset successfully for: ' + username);
        return { 
          success: true, 
          message: 'Password reset successful! Check your email for the new password.' 
        };
      }
    }
    return { success: false, message: 'Username and email combination not found' };
  } catch (error) {
    Logger.log('Password reset error: ' + error.toString());
    return { success: false, message: 'Error resetting password: ' + error.toString() };
  }
}

// Verify credentials - SIMPLIFIED VERSION
function verifyCredentials(username, password) {
  try {
    Logger.log('Login attempt for: ' + username);
    
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === username.trim()) {
        if (data[i][3] === password) {
          Logger.log('Login successful for: ' + username);
          return { 
            success: true, 
            message: 'Login successful',
            userData: {
              username: data[i][1],
              email: data[i][2],
              name: data[i][4]
            }
          };
        } else {
          return { success: false, message: 'Invalid password' };
        }
      }
    }
    return { success: false, message: 'Username not found' };
  } catch (error) {
    Logger.log('Login error: ' + error.toString());
    return { success: false, message: 'Error during login: ' + error.toString() };
  }
}

// Update user data
function updateUserData(username, updatedData) {
  try {
    var sheet = initializeSystem();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === username.trim()) {
        var row = i + 1; // +1 because array is 0-indexed but sheet rows start at 1
        
        Logger.log('Updating user: ' + username);
        Logger.log('Update data: ' + JSON.stringify(updatedData));
        
        // Update fields - using exact column indices
        if (updatedData.name !== undefined && updatedData.name !== '') 
          sheet.getRange(row, 5).setValue(updatedData.name);
        if (updatedData.address !== undefined && updatedData.address !== '') 
          sheet.getRange(row, 6).setValue(updatedData.address);
        if (updatedData.phone !== undefined && updatedData.phone !== '') 
          sheet.getRange(row, 7).setValue(updatedData.phone);
        if (updatedData.dob !== undefined && updatedData.dob !== '') 
          sheet.getRange(row, 8).setValue(updatedData.dob);
        if (updatedData.gender !== undefined && updatedData.gender !== '') 
          sheet.getRange(row, 9).setValue(updatedData.gender);
        if (updatedData.nationality !== undefined && updatedData.nationality !== '') 
          sheet.getRange(row, 10).setValue(updatedData.nationality);
        
        // Update last modified timestamp
        sheet.getRange(row, 12).setValue(new Date());
        
        return { success: true, message: 'Profile updated successfully!' };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (error) {
    Logger.log('Update error: ' + error.toString());
    return { success: false, message: 'Error updating profile: ' + error.toString() };
  }
}

// Helper function to format dates for display
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  try {
    var date = new Date(dateValue);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return dateValue.toString();
  }
}

// Send password change confirmation email
function sendPasswordChangeEmail(email, username, name) {
  try {
    var subject = CONFIG.SERVICE_NAME + ' - Password Changed Successfully';
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #072033;">Password Update Confirmation</h2>
        <p>Hello ${name},</p>
        <p>Your password for E-Huduma account <strong>${username}</strong> has been successfully changed.</p>
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
          <p><strong>Username:</strong> ${username}</p>
          <p><strong>Time of change:</strong> ${new Date().toLocaleString()}</p>
        </div>
        <p style="color: #666;">If you did not make this change, please contact support immediately.</p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 30px 0;">
        <p style="color: #999; font-size: 12px;">Security notice - E-Huduma</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (error) {
    Logger.log('Password change email error: ' + error.toString());
  }
}

// Send password reset email
function sendPasswordResetEmail(email, username, newPassword, name) {
  try {
    var subject = CONFIG.SERVICE_NAME + ' - Password Reset';
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #072033;">Password Reset Complete</h2>
        <p>Hello ${name},</p>
        <p>Your password for E-Huduma account <strong>${username}</strong> has been reset.</p>
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
          <p><strong>Username:</strong> ${username}</p>
          <p><strong>New Password:</strong> ${newPassword}</p>
          <p><strong>Time of reset:</strong> ${new Date().toLocaleString()}</p>
        </div>
        <p style="color: #666; font-size: 14px;">
          <strong>Important:</strong> Please login and change your password immediately for security.
        </p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 30px 0;">
        <p style="color: #999; font-size: 12px;">Security notice - E-Huduma</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (error) {
    Logger.log('Password reset email error: ' + error.toString());
  }
}

// Web app endpoints
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}