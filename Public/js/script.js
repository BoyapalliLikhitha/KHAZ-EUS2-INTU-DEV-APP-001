// Function to log messages
const closeBtn=document.querySelector('.close');
closeBtn.addEventListener('click',()=>{hideDialog('messageDialog')});

const themeToggle = document.getElementById('themeToggle');

themeToggle.addEventListener('change', () => {
    if (themeToggle.checked) {
        document.body.classList.add('light-mode'); // Switch to light mode
    } else {
        document.body.classList.remove('light-mode'); // Switch back to dark mode
    }
});

function log(message, type = 'info') {
    const logEntry = `[${new Date().toISOString()}] ${type.toUpperCase()}: ${message}`;
    console.log(logEntry);

    // Send log to the server
    fetch(`${ServerUrl}/log`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ message: logEntry, type })
    })
    .catch(error => console.error('Error sending log to server:', error));
}

// Function to show dialog by ID and optionally set message
function showDialog(elementId, message) {
    const dialog = document.getElementById(elementId);
    if (dialog) {
        dialog.style.display = 'block'; // Display the dialog
        log(`Dialog "${elementId}" shown`);

        // Set message if provided
        if (message) {
            const messageElement = dialog.querySelector('.message');
            if (messageElement) {
                messageElement.textContent = message;
                log(`Message in dialog "${elementId}" set to: "${message}"`);
            }
        }
    }
}

// Function to hide dialog by ID
function hideDialog(elementId) {
    const dialog = document.getElementById(elementId);
    if (dialog) {
        dialog.style.display = 'none'; // Hide the dialog
        log(`Dialog "${elementId}" hidden`);
    }
}
function showMessage(message, type) {
    const messageIcon = document.getElementById('messageIcon');
    const messageText = document.getElementById('messageText');
    
    if (type === 'success') {
        messageIcon.className = 'icon success';
        messageIcon.innerHTML = '✔️';
    } else if (type === 'error') {
        messageIcon.className = 'icon error';
        messageIcon.innerHTML = '❌';
    }

    messageText.textContent = message;
    showDialog('messageDialog');
}


// Fetch access token and recovery key on page load
let authforApp = ''
let authforSp =''
let loggedUserEmail = ''

const ServerUrl = 'https://khaz-eus2-intu-dev-app-001.azurewebsites.net';
const siteUrl="https://stefaninidemo1.sharepoint.com/sites/IntuneUtility";
const listName="DeviceActionsLog";

/*function fetchUserEmail() {
    fetch(`${ServerUrl}/.auth/me`, {
        method: 'GET',
        credentials: 'include' // This allows cookies to be sent with the request
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Failed to retrieve user details');
        }
        return response.json();
    })
    .then(data => {
        if (data && data.length > 0) {
            const user = data[0];
            
            // Find the email claim and set the global variable
            loggedUserEmail = user.user_claims.find(claim => claim.typ === "preferred_username")?.val || '';
            return loggedUserEmail;
            log(`Logged in as: ${loggedUserEmail}`);
        } else {
            log("No user details found.");
            loggedUserEmail = ''; // Clear email on failed fetch
        }
    })
    .catch(error => {
        console.error("Error fetching user details:", error);
        loggedUserEmail = ''; // Clear email on error
    });
}
*/

async function fetchUserEmail() {
    try {
        const response = await fetch(`${ServerUrl}/.auth/me`, {
            method: 'GET',
            credentials: 'include' // This allows cookies to be sent with the request
        });

        if (!response.ok) {
            throw new Error('Failed to retrieve user details');
        }

        const data = await response.json();

        if (data && data.length > 0) {
            const user = data[0];
            const loggedUserEmail = user.user_claims.find(claim => claim.typ === "preferred_username")?.val || '';
            log(`Logged in as: ${loggedUserEmail}`);
            return loggedUserEmail; // Return email
        } else {
            log("No user details found.");
            return ''; // Return empty email
        }
    } catch (error) {
        console.error("Error fetching user details:", error);
        return ''; // Return empty email on error
    }
}

// Usage
fetchUserEmail().then(email => {
    loggedUserEmail = email;
    console.log("Fetched email:", loggedUserEmail);
});

if (!authforApp ) {
    
    fetch(`${ServerUrl}/getAccessToken`)
        .then(response => {
            if (!response.ok) {
                log(response)
                log('Failed to get app token', 'error');
                throw new Error('Failed to get app token');
            }
            return response.json();
        })
        .then(data => {
            authforApp = data.app_token;
            authforSp= data.SP_Token;
            
            log('Successfully fetched access token ');
        })
        .catch(error => {
            log(`Error fetching tokens: ${error.message}`, 'error');
        });
}



// Function to send the action data to SharePoint
async function sendActionToSharePoint(ticketNo, userEmail, deviceName, actionPerformed,loggedUserEmail) {
    
    const actionData = {
        Ticket: ticketNo,
        User: userEmail,
        DeviceSelected: deviceName,
        ActionPerformed: actionPerformed,
        LoggedUser:loggedUserEmail,
        Timestamp: new Date().toISOString()
    };

    const requestBody = {
        __metadata: { "type": 'SP.Data.DeviceActionsLogListItem' },
        "Title": actionData.Ticket,
        "User": actionData.User,
        "DeviceSelected": actionData.DeviceSelected,
        "ActionPerformed": actionData.ActionPerformed,
        "LoggedUser":actionData.LoggedUser,
        "Timestamp": actionData.Timestamp
    };

    try {
        const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${authforSp}`,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
            },
            body: JSON.stringify(requestBody)
        });

        // Log raw response for debugging
        const responseText = await response.text();  // Read response as text

        // Check if the response is valid JSON
        try {
            const data = JSON.parse(responseText);
            
            if (data.d && data.d.Id) {
                log('Action successfully logged to SharePoint');
            } else {
               log('Failed to log action to SharePoint');
            }
        } catch (jsonError) {
            log('Failed to parse JSON response:', jsonError);
        }
    } catch (error) {
        log('Error sending data to SharePoint:', error);
    }
}


// Function to check if email is valid
function isEmailValid(email) {
    const emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
    return emailRegex.test(email);
}

// Function to enable/disable buttons based on input status
function updateButtonStatus() {
    const ticketNoInput = document.getElementById('ticketNo');
    const userEmailInput = document.getElementById('userEmail');
    const submitBtn = document.getElementById('submitBtn');

    if (ticketNoInput && userEmailInput && submitBtn) {
        //submitBtn.disabled = !(ticketNoInput.value.trim() !== '' && isEmailValid(userEmailInput.value.trim()));
        if(isEmailValid(userEmailInput.value.trim())){
        log(`Inputs Provided : [${ticketNoInput.value},${userEmailInput.value}]`);
        }
    }
}

// Add event listeners to input fields to update button status
document.addEventListener('DOMContentLoaded', () => {
    log('Document loaded, setting up event listeners');
    const ticketNoInput = document.getElementById('ticketNo');
    const userEmailInput = document.getElementById('userEmail');
    const submitBtn = document.getElementById('submitBtn');

    if (ticketNoInput && userEmailInput && submitBtn) {
        ticketNoInput.addEventListener('input', updateButtonStatus);
        userEmailInput.addEventListener('input', updateButtonStatus);
    }
});

// Event listener for the Submit button
document.getElementById('submitBtn').addEventListener('click', () => {
    log('Submit button clicked');
    
    const ticketNoInput = document.getElementById('ticketNo');
    const userEmailInput = document.getElementById('userEmail');
    const deviceListSelect = document.getElementById('deviceList');

    if (ticketNoInput.value.trim() === '' || !isEmailValid(userEmailInput.value.trim())) {
        log('Invalid input: Ticket No or Email is not valid', 'warning');
        showMessage('Please enter valid Ticket No and User Email.','error');
        return;
    }

    hideDialog('inputDialog');
    showDialog('loadingSpinner');
    log('Fetching user devices');

    const userEmail = userEmailInput.value.trim();

    fetch(`${ServerUrl}/getUserDevices`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${authforApp}`
        },
        body: JSON.stringify({ userEmail })
    })
        .then(response => {
            if (!response.ok) {
                log('Failed to fetch user devices', 'error');
                throw new Error('Failed to fetch user devices');
            }
            return response.json();
        })
        .then(devices => {
            log('Successfully fetched user devices');
            deviceListSelect.innerHTML = '<option value="" selected disabled>Select a Device</option>';
            devices.forEach(device => {
                const option = document.createElement('option');
                option.value = device.id;
                option.textContent = device.deviceName;
                deviceListSelect.appendChild(option);
            });
            hideDialog('loadingSpinner');
            showDialog('deviceDialogOverlay');
        })
        .catch(error => {
            log(`Error fetching user devices: ${error.message}`, 'error');
            showMessage('Devices do not exist (or) Failed to fetch user devices', 'error');
            hideDialog('loadingSpinner');
            
        });
});


const TicketNo =document.getElementById('ticketNo')
const userEmail =document.getElementById('userEmail')

// Event listener for device selection
document.getElementById('deviceList').addEventListener('change', () => {
    log('Device selected from list');
    
    const selectedDeviceId = document.getElementById('deviceList').value;

    fetch(`${ServerUrl}/getDeviceInfo`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${authforApp}`
        },
        body: JSON.stringify({ deviceId: selectedDeviceId })
    })
        .then(response => {
            if (!response.ok) {
                log('Failed to fetch device information', 'error');
                throw new Error('Failed to fetch device information');
            }
            return response.json();
        })
        .then(deviceInfo => {
            log(`Successfully fetched device information for Device ID: ${selectedDeviceId}`);
            document.getElementById('deviceName').textContent = deviceInfo.deviceName;
            document.getElementById('azureId').textContent = deviceInfo.azureADDeviceId;
            document.getElementById('id').textContent = deviceInfo.id;
            document.getElementById('ownership').textContent = deviceInfo.managedDeviceOwnerType;
            document.getElementById('phoneNumber').textContent = deviceInfo.phoneNumber || 'N/A';
            document.getElementById('deviceManufacturer').textContent = deviceInfo.manufacturer;
            document.getElementById('primaryUser').textContent = deviceInfo.userPrincipalName;
            document.getElementById('enrolledDate').textContent = new Date(deviceInfo.enrolledDateTime).toLocaleString();
            document.getElementById('compliance').textContent = deviceInfo.complianceState;
            document.getElementById('operatingSystem').textContent = deviceInfo.operatingSystem;
            document.getElementById('deviceModel').textContent = deviceInfo.model;
            document.getElementById('lastCheckInTime').textContent = new Date(deviceInfo.lastSyncDateTime).toLocaleString();
            document.getElementById('osVersion').textContent = deviceInfo.osVersion;

            showDialog('deviceDialogOverlay');
        })
        .catch(error => {
            log(`Error fetching device information: ${error.message}`, 'error');
            showMessage('Failed to fetch device information', 'error');
        });
});

// Event listener for Sync button
document.getElementById('sync').addEventListener('click', () => {
    initiateSync();
});

// Function to initiate Sync
function initiateSync() {
    log('Sync action initiated');
    
    const selectedDeviceId = document.getElementById('id').textContent;
    const selectedDeviceName = document.getElementById('deviceName').textContent;
    const syncEndpoint = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${selectedDeviceId}/microsoft.graph.syncDevice`;
    showDialog('loadingSpinner');
    const headers = new Headers({
        'Authorization': `Bearer ${authforApp}`
    });

    const requestOptions = {
        method: 'POST',
        headers: headers,
        redirect: 'follow'
    };

    fetch(syncEndpoint, requestOptions)
        .then(response => {
            if (response.ok) {
                log(`Sync initiated for Device: ${selectedDeviceName}`);
                hideDialog('loadingSpinner');
                showMessage(`Machine ${selectedDeviceName} will be synced shortly.`, 'success');
                sendActionToSharePoint(TicketNo.value.trim(),userEmail.value.trim(),selectedDeviceName,"Sync",loggedUserEmail)
            } else {
                log(`Failed to sync Machine ${selectedDeviceName}`, 'error');
                hideDialog('loadingSpinner');
                showMessage(`Failed to sync Machine ${selectedDeviceName}. Please try again after some time.`, 'error');
            }
        })
        .catch(error => {
            log(`Error during sync action: ${error.message}`, 'error');
            hideDialog('loadingSpinner');
            showMessage('Something went wrong during sync action. Please try again after some time.', 'error');
        });
}

// Event listener for Reboot button
document.getElementById('reboot').addEventListener('click', () => {
    log('Reboot button clicked');
    
    const selectedDeviceId = document.getElementById('id').textContent;
    const selectedDeviceName = document.getElementById('deviceName').textContent;
    const rebootEndpoint = `https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${selectedDeviceId}/microsoft.graph.rebootNow`;
    showDialog('loadingSpinner');
    const headers = new Headers({
        'Authorization': `Bearer ${authforApp}`
    });

    const requestOptions = {
        method: 'POST',
        headers: headers,
        redirect: 'follow'
    };

    fetch(rebootEndpoint, requestOptions)
        .then(response => {
            if (response.ok) {
                log(`Reboot initiated for Device: ${selectedDeviceName}`);
                hideDialog('loadingSpinner');
                showMessage(`Machine ${selectedDeviceName} will be rebooted shortly. Sync will initiate immediately.`, 'success');
                sendActionToSharePoint(TicketNo.value.trim(),userEmail.value.trim(),selectedDeviceName,"Reboot",loggedUserEmail)
                // Immediately initiate sync after reboot request
                setTimeout(initiateSync, 2000);
                
            } else {
                log(`Failed to reboot Machine ${selectedDeviceName}`, 'error');
                hideDialog('loadingSpinner');
                showMessage(`Failed to reboot Machine ${selectedDeviceName}. Please try again after some time.`, 'error');
            }
        })
        .catch(error => {
            log(`Error during reboot action: ${error.message}`, 'error');
            hideDialog('loadingSpinner');
            showMessage('Something went wrong during reboot action. Please try again after some time.', 'error');
        });
});

document.getElementById('recoveryKey').addEventListener('click',async()=>{
    log('Recoverykey button clicked');

    const selectedDeviceId = document.getElementById('azureId').textContent;
    const selectedDeviceName = document.getElementById('deviceName').textContent;
    
    // Show loading spinner
    showDialog('loadingSpinner');

    try {   

        // Define the Microsoft Graph API URL for BitLocker recovery keys
        const recoveryKeyEndpoint = `https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys?$filter=deviceId eq '${selectedDeviceId}'&$orderby=createdDateTime desc`;

        // Create headers with the access token
        const headers = new Headers({
            'Authorization': `Bearer ${authforApp}`
        });

        // Create the fetch request options
        const requestOptions = {
            method: 'GET',
            headers: headers,
            redirect: 'follow'
        };

        // Perform the BitLocker recovery key retrieval operation
        const response = await fetch(recoveryKeyEndpoint, requestOptions);
        const data =  await response.json();

        if (data.value.length > 0) {
            const recoveryKeyId = data.value[0].id;
            const recoveryKeyDetailsEndpoint = `https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/${recoveryKeyId}?$select=key`;

            // Fetch the BitLocker recovery key details
            const keyResponse =  await fetch(recoveryKeyDetailsEndpoint, requestOptions);
            const recoveryKeyData = await keyResponse.json();

            const recoveryKey = recoveryKeyData.key;
            
             // Display the BitLocker recovery key/'
            log(`Recovery Key for ${selectedDeviceName} is: ${recoveryKey}`);
            hideDialog('loadingSpinner');
            showMessage(`${recoveryKey}`,'success');
            sendActionToSharePoint(TicketNo.value.trim(),userEmail.value.trim(),selectedDeviceName,"RecoveryKey",loggedUserEmail)
        } else {
            log(`Machine ${selectedDeviceName} does not have any Recovery key attached to it`);
            hideDialog('loadingSpinner');
            showMessage(`Machine ${selectedDeviceName} does not have any Recovery key attached to it...`,'success');
        }
    
    }
 catch (error) {
    hideDialog('loadingSpinner');
    log('Error obtaining BitLocker recovery key. Please try again later.','error')
    console.error('Error obtaining BitLocker recovery key:', error);
    showMessage('Error obtaining BitLocker recovery key. Please try again later.','error');
}
});


// Event listener for "Fetch Credentials" button
document.getElementById('localAdminPassword').addEventListener('click', () => {
    log('Fetch Credentials button clicked');
    
    const selectedDeviceId = document.getElementById('azureId').textContent;
    const selectedDeviceName = document.getElementById('deviceName').textContent;

    const credentialsEndpoint = `https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials/${selectedDeviceId}?$select=credentials`;
    showDialog('loadingSpinner');
    
    const headers = new Headers({
        'Authorization': `Bearer ${authforApp}`
    });

    const requestOptions = {
        method: 'GET', // Using GET method for fetching data
        headers: headers,
        redirect: 'follow'
    };

    fetch(credentialsEndpoint, requestOptions)
        .then(response => response.json())
        .then(data => {
            if (data.credentials && data.credentials.length>0) {
                // Decode the Base64-encoded password
                const credential = data.credentials[0];
                log(credential)
                const decodedPassword = atob(credential.passwordBase64);
                
                log(`Credentials fetched for Device ID: ${selectedDeviceId}`);
                hideDialog('loadingSpinner');
                showMessage(`Device password: ${decodedPassword}`, 'success');
                sendActionToSharePoint(TicketNo.value.trim(),userEmail.value.trim(),selectedDeviceName,"LAPS",loggedUserEmail)
            } else {
                log(`Failed to fetch credentials for Device ID: ${selectedDeviceId}`, 'error');
                hideDialog('loadingSpinner');
                showMessage(`No credentials available for the device.`, 'error');
            }
        })
        .catch(error => {
            log(`Error fetching credentials: ${error.message}`, 'error');
            hideDialog('loadingSpinner');
            showMessage('No credentials available for the device', 'error');
        });
});


cancelButton=document.getElementById('logout')
document.getElementById('reset').addEventListener('click', () => {
    location.reload();  // This will refresh the page
});
cancelButton.addEventListener('click', () => {
    //Show a confirmation dialog
   var result = confirm('Are you sure you want to log out?');
    if (result) {
        // User clicked 'OK' (Yes), proceed with the logout action
        
        window.location.href = `${ServerUrl}/.auth/logout`;
        log(s)
    } 
});
