// index.js
require('dotenv').config({ path: '.env' });
require('dotenv').config({ path: '.env.local' });
const https = require('https');
const http = require('http');
const querystring = require('querystring');

// Environment variables
const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

if (!CLIENT_ID || !TENANT_ID || !CLIENT_SECRET) {
  console.error('Missing environment variables. Please check your .env and .env.local files.');
  process.exit(1);
}

// Function to obtain access token from Azure AD
function getAccessToken(callback) {
  const postData = querystring.stringify({
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: CLIENT_SECRET,
    grant_type: 'client_credentials',
  });

  const options = {
    hostname: 'login.microsoftonline.com',
    path: `/${TENANT_ID}/oauth2/v2.0/token`,
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(postData),
    },
  };

  const req = https.request(options, (res) => {
    let body = '';
    res.on('data', (chunk) => (body += chunk));
    res.on('end', () => {
      if (res.statusCode === 200) {
        const response = JSON.parse(body);
        callback(null, response.access_token);
      } else {
        const error = new Error(`Token request failed with status ${res.statusCode}: ${body}`);
        console.error('Error obtaining access token:', error);
        callback(error);
      }
    });
  });

  req.on('error', (err) => {
    console.error('Error in access token request:', err);
    callback(err);
  });
  req.write(postData);
  req.end();
}

// Function to fetch users from Microsoft Graph API
function getUsers(accessToken, callback) {
  const options = {
    hostname: 'graph.microsoft.com',
    path: '/v1.0/users',
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  };

  const req = https.request(options, (res) => {
    let body = '';
    res.on('data', (chunk) => (body += chunk));
    res.on('end', () => {
      if (res.statusCode === 200) {
        const response = JSON.parse(body);
        callback(null, response);
      } else {
        const error = new Error(`API request failed with status ${res.statusCode}: ${body}`);
        console.error('Error fetching users:', error);
        callback(error);
      }
    });
  });

  req.on('error', (err) => {
    console.error('Error in user fetch request:', err);
    callback(err);
  });
  req.end();
}

// Create HTTP server
const server = http.createServer((req, res) => {
  // Set common CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.url === '/users') {
    if (req.method === 'OPTIONS') {
      // Handle preflight OPTIONS request
      res.writeHead(204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      });
      res.end();
    } else if (req.method === 'GET') {
      getAccessToken((err, token) => {
        if (err) {
          console.error('Error obtaining access token from Azure AD:', err);
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end(`Error obtaining access token: ${err.message}`);
          return;
        }
        getUsers(token, (err, data) => {
          if (err) {
            console.error('Error fetching users from Microsoft Graph API:', err);
            res.writeHead(500, { 'Content-Type': 'text/plain' });
            res.end(`Error fetching users: ${err.message}`);
            return;
          }
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(data));
        });
      });
    } else {
      console.error('Method Not Allowed:', req.method);
      res.writeHead(405, {
        'Content-Type': 'text/plain',
        'Access-Control-Allow-Methods': 'GET, OPTIONS',
      });
      res.end('Method Not Allowed');
    }
  } else {
    console.error('Not Found:', req.url);
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found');
  }
});

// Start the server
const PORT = process.env.PORT || 3001;
server.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
