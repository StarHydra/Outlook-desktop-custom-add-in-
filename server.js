const https = require('https');
const fs = require('fs');

//Add your SendGrid API key here
const SENDGRID_API_KEY = '';

const DEFAULT_RECIPIENT_EMAIL = 'abc@xyz.com';//Replace with your default recipient email
const SENDER_EMAIL = 'xyz@abc.com';//Replace with your verified sender email in SendGrid
const SENDER_NAME = 'Text Forwarder';

//DEBUG MODE: Set to true to send a simple test message instead of actual content
const DEBUG_MODE = false;

//Loading SSL certificates
const options = {
  key: fs.readFileSync('cert.key'),
  cert: fs.readFileSync('cert.crt')
};

//Creating HTTPS server
const server = https.createServer(options, (req, res) => {
  //Setting CORS headers to allow ngrok
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, ngrok-skip-browser-warning');

  //Handle preflight request
  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  //Handling POST request to send email
  if (req.method === 'POST' && req.url === '/send-email') {
    let body = '';

    req.on('data', chunk => {
      body += chunk.toString();
    });

    req.on('end', () => {
      try {
        const data = JSON.parse(body);
        const { text, subject, recipient } = data;
        
        //Logging received data for debugging
        console.log('Received request:');
        console.log('- Subject:', subject);
        console.log('- Recipient:', recipient);
        console.log('- Text length:', text ? text.length : 0);
        console.log('- Text preview (first 100 chars):', text ? text.substring(0, 100) : 'null');
        console.log('- Text has special chars:', /[^\x20-\x7E\n\r\t]/.test(text));
        
        //Validating required fields
        if (!text || !subject) {
          console.error('Missing required fields: text or subject');
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ 
            success: false, 
            message: 'Missing required fields: text and subject are required' 
          }));
          return;
        }
        
        //Cleaning the text to remove all problematic characters
        let cleanText = text
          //Normalizing unicode characters
          .normalize('NFKD')
          //Removing null and other dangerous chars
          .replace(/\u0000/g, '')
          //Removing zero-width characters
          .replace(/[\u200B-\u200D\uFEFF]/g, '')
          //Removing all control characters except newline, tab, carriage return
          .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
          //Normalizing line endings
          .replace(/\r\n/g, '\n')
          .replace(/\r/g, '\n')
          //Removing any non-ASCII characters that might cause issues
          //Only letters, numbers, punctuation, spaces, newlines and tabs kept
          .replace(/[^\x20-\x7E\n\t]/g, '')
          .trim();
        
        //Cleaning subject the same way
        const cleanSubject = (subject || 'Forwarded Content')
          .normalize('NFKD')
          .replace(/\u0000/g, '')
          .replace(/[\u200B-\u200D\uFEFF]/g, '')
          .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
          .replace(/[^\x20-\x7E ]/g, '')
          .trim();
        
        //Validating cleaned text
        if (!cleanText) {
          console.error('Text became empty after cleaning');
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ 
            success: false, 
            message: 'Email content could not be processed' 
          }));
          return;
        }
        
        //Using provided recipient or default
        const recipientEmail = recipient || DEFAULT_RECIPIENT_EMAIL;
        
        console.log('Sending email to:', recipientEmail);
        console.log('Clean subject:', cleanSubject);
        console.log('Clean text length:', cleanText.length);
        console.log('Clean text preview:', cleanText.substring(0, 100));
        console.log('Clean text has special chars:', /[^\x20-\x7E\n\t]/.test(cleanText));

        //Preparing SendGrid email data
        let emailPayload;
        
        if (DEBUG_MODE) {
          //Sending a simple test message
          console.log('DEBUG MODE: Sending simple test message instead of actual content');
          emailPayload = {
            personalizations: [{
              to: [{ email: recipientEmail }]
            }],
            from: {
              email: SENDER_EMAIL,
              name: SENDER_NAME
            },
            subject: "Test Email from Outlook Add-in",
            content: [{
              type: 'text/plain',
              value: 'This is a simple test message to verify SendGrid is working correctly.'
            }]
          };
        } else {
          //Sending actual email content
          emailPayload = {
            personalizations: [{
              to: [{ email: recipientEmail }]
            }],
            from: {
              email: SENDER_EMAIL,
              name: SENDER_NAME
            },
            subject: cleanSubject,
            content: [{
              type: 'text/plain',
              value: cleanText
            }]
          };
        }
        
        const emailData = JSON.stringify(emailPayload);
        
        //Logging the email data being sent (first 300 chars to see structure)
        console.log('Email data being sent to SendGrid:');
        console.log(emailData.substring(0, 300));
        console.log('Email data length:', emailData.length);

        //Sending to SendGrid
        const options = {
          hostname: 'api.sendgrid.com',
          port: 443,
          path: '/v3/mail/send',
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${SENDGRID_API_KEY}`,
            'Content-Type': 'application/json',
            'Content-Length': emailData.length
          }
        };

        const sendGridReq = https.request(options, (sendGridRes) => {
          let responseData = '';

          sendGridRes.on('data', (chunk) => {
            responseData += chunk;
          });

          sendGridRes.on('end', () => {
            console.log('SendGrid response status:', sendGridRes.statusCode);
            console.log('SendGrid response body:', responseData);
            
            if (sendGridRes.statusCode === 202 || sendGridRes.statusCode === 200) {
              console.log('✓ Email sent successfully');
              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ success: true, message: 'Email sent successfully' }));
            } else {
              console.error('✗ SendGrid rejected the email');
              res.writeHead(sendGridRes.statusCode, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ 
                success: false, 
                message: 'Failed to send email', 
                error: responseData,
                statusCode: sendGridRes.statusCode
              }));
            }
          });
        });

        sendGridReq.on('error', (error) => {
          console.error('SendGrid request error:', error);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, message: error.message }));
        });

        sendGridReq.write(emailData);
        sendGridReq.end();

      } catch (error) {
        console.error('Error processing request:', error);
        console.error('Request body:', body);
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ 
          success: false, 
          message: 'Invalid request: ' + error.message 
        }));
      }
    });
  } else {
    res.writeHead(404);
    res.end('Not Found');
  }
});

const PORT = 3001;
server.listen(PORT, () => {
  console.log(`Backend server running on https://localhost:${PORT}`);
  console.log('Ready to forward emails to SendGrid!');
});