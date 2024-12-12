import express from 'express';
import fetch from 'node-fetch';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import axios from 'axios';
import bodyParser from 'body-parser';
import winston from 'winston';
import 'winston-daily-rotate-file';
import fs from 'fs';


dotenv.config();

const app = express();

// Middleware
app.use(bodyParser.json());
app.use(cors());
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const logDirectory = path.join(__dirname, 'public/logs');
if (!fs.existsSync(logDirectory)) {
    fs.mkdirSync(logDirectory, { recursive: true });
}

// Setup Winston logger with daily log rotation
const logger = winston.createLogger({
    level: 'info',
    format: winston.format.combine(
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
        winston.format.printf(({ timestamp, level, message }) => `${timestamp} [${level.toUpperCase()}]: ${message}`)
    ),
    transports: [
        new winston.transports.Console(),
        new winston.transports.DailyRotateFile({
            filename: path.join(logDirectory, '%DATE%.log'),
            datePattern: 'YYYY-MM-DD',
            maxSize: '20m',
            maxFiles: '90d',
            zippedArchive: true,
        }),
    ]
});

// Azure AD configuration
const clientId = "e752e188-a26b-4cd2-adef-06d48013161a";
const clientSecret = "8dr8Q~nBH_6MDcZLKqzFiocmwUmOEp2E6CW9GcWa";
const tenantId = "3ce34e42-c07d-47bb-b72a-4ce606de6b88";
const sp_clientId = "b87e48d6-2c35-4937-aef2-ff5c3f68b01d";
const sp_clientSecret = "r.H8Q~ltGKpUeoUIwgFLbNrrCcT6WjbQPT5DIa3W";
const sp_tenantId ="fdc223cd-8687-48e5-b9e8-ad52f8adbdaa";
const User = "nbhurli@stefaninidemo1.onmicrosoft.com";
const password ="Stefanini37@123!";


app.use(express.static(path.join(__dirname, 'public')));


// Log out and redirect (since no session, just log and redirect)
app.get('/logout', (req, res) => {
    logger.info('User session ended');
    res.redirect('/');
});

// Endpoint to get access tokens
app.get('/getAccessToken', async (req, res) => {
    try {
        const appAuthToken = await authForApp(clientId, clientSecret, tenantId);
        const SPToken =await SP_auth(sp_clientId, sp_clientSecret,sp_tenantId);
        

        logger.info('Successfully obtained access tokens');
        res.json({
            app_token: appAuthToken,
            SP_Token: SPToken
            
        });
    } catch (error) {
        logger.error('Error obtaining access tokens:', error);
        res.status(500).json({ error: 'Failed to obtain access tokens' });
    }
});

// Function to get Intune API access token
async function authForApp(clientId, clientSecret, tenantId) {
    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const authRequestBody = {
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default',
        client_id: clientId,
        client_secret: clientSecret,
    };

    try {
        const response = await fetch(authUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams(authRequestBody).toString(),
        });

        const authData = await response.json();
        if (authData.access_token) {
            return authData.access_token;
        }
        throw new Error('No access token found');
    } catch (error) {
        logger.error('Error obtaining app token:', error);
        throw new Error('Error obtaining app token');
    }
}
async function SP_auth(sp_clientId, sp_clientSecret, sp_tenantId) {
    const authUrl = `https://login.microsoftonline.com/${sp_tenantId}/oauth2/v2.0/token`;
    const authRequestBody = {
        grant_type: 'password',
        scope: 'https://stefaninidemo1.sharepoint.com/.default',
        client_id: sp_clientId,
        client_secret: sp_clientSecret,
        username:User,
        password:password   
    };

    try {
        const response = await fetch(authUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams(authRequestBody).toString(),
        });

        const authData = await response.json();
        if (authData.access_token) {
            return authData.access_token;
        }
        throw new Error('No access token found');
    } catch (error) {
        console.error('Error obtaining app token:', error);
        throw new Error('Error obtaining app token', error);
    }
}

// Example API endpoint to fetch user devices after authentication (SSO required)
app.post('/getUserDevices', async (req, res) => {
    const { userEmail } = req.body;

    try {
        const token = await authForApp(clientId, clientSecret, tenantId);
        const response = await axios.get(`https://graph.microsoft.com/v1.0/users/${userEmail}/managedDevices`, {
            headers: { Authorization: `Bearer ${token}` },
        });

        // Check if devices are null, undefined, or an empty array
        if (!response.data.value || response.data.value.length === 0) {
            logger.info(`No devices found for user: ${userEmail}`);
            return res.status(404).send('No devices are found');
        }

        logger.info(`Fetched user devices successfully for: ${userEmail}`);
        res.json(response.data.value);
    } catch (error) {
        logger.error('Error fetching user devices:', error);
        res.status(500).send('Failed to fetch user devices');
    }
});

// Example API endpoint to fetch device information (SSO required)
app.post('/getDeviceInfo', async (req, res) => {
    const { deviceId } = req.body;

    try {
        const token = await authForApp(clientId, clientSecret, tenantId);
        const response = await axios.get(`https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${deviceId}`, {
            headers: { Authorization: `Bearer ${token}` },
        });

        logger.info(`Fetched device information successfully for: ${deviceId}`);
        res.json(response.data);
    } catch (error) {
        logger.error('Error fetching device information:', error);
        res.status(500).send('Failed to fetch device information');
    }
});
app.post('/log', (req, res) => {
    const { message, type } = req.body;
    if (message && type) {
        logger[type](message); // Use the appropriate logging level
        res.status(200).send('Log received');
    } else {
        res.status(400).send('Invalid log data');
    }
});

// Start the server
const PORT = process.env.PORT || 7767; // Use Azure's provided PORT or fallback to 7767 for local testing
app.listen(PORT, () => {
    logger.info(`Server running on port ${PORT}`);
});

