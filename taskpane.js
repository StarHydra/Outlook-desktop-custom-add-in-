//const BACKEND_URL = 'https://localhost:3001/send-email';
const BACKEND_URL = 'https://contrastable-overaffected-irma.ngrok-free.dev/send-email';
const DEFAULT_RECIPIENT = 'tairsxc@gmail.com';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    //Setting up event listeners
    document.getElementById('entireEmailBtn').onclick = () => switchMode('entire');
    document.getElementById('selectedTextBtn').onclick = () => switchMode('selected');
    document.getElementById('sendEntireBtn').onclick = sendEntireEmail;
    document.getElementById('sendSelectedBtn').onclick = sendSelectedText;
    
    //Recipient radio button listeners
    const radioButtons = document.querySelectorAll('input[name="recipient"]');
    radioButtons.forEach(radio => {
      radio.addEventListener('change', handleRecipientChange);
    });
    
    //Loading email preview
    loadEmailPreview();
    
    //Trying to auto-populate selected text
    // tryGetSelectedText();
  }
});

//Function to handle recipient selection change
function handleRecipientChange(event) {
  const customSection = document.getElementById('customEmailSection');
  if (event.target.value === 'custom') {
    customSection.style.display = 'block';
    document.getElementById('customEmail').focus();
  } else {
    customSection.style.display = 'none';
  }
}

//Function to get the recipient email address
function getRecipientEmail() {
  const selectedRecipient = document.querySelector('input[name="recipient"]:checked').value;
  if (selectedRecipient === 'default') {
    return DEFAULT_RECIPIENT;
  } else {
    const customEmail = document.getElementById('customEmail').value.trim();
    if (!customEmail) {
      throw new Error('Please enter a custom email address');
    }
    //Validating email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(customEmail)) {
      throw new Error('Please enter a valid email address');
    }
    return customEmail;
  }
}

//Function to switch between modes
function switchMode(mode) {
  const entireBtn = document.getElementById('entireEmailBtn');
  const selectedBtn = document.getElementById('selectedTextBtn');
  const entireSection = document.getElementById('entireEmailSection');
  const selectedSection = document.getElementById('selectedTextSection');
  
  if (mode === 'entire') {
    entireBtn.classList.add('active');
    selectedBtn.classList.remove('active');
    entireSection.style.display = 'block';
    selectedSection.style.display = 'none';
  } else {
    entireBtn.classList.remove('active');
    selectedBtn.classList.add('active');
    entireSection.style.display = 'none';
    selectedSection.style.display = 'block';
    // tryGetSelectedText(); //Uncomment to auto-fetch selected text
  }
}

//Function to try to get selected text from email
// function tryGetSelectedText() {
//   if (Office.context.mailbox.item.getSelectedDataAsync) {
//     Office.context.mailbox.item.getSelectedDataAsync(
//       Office.CoercionType.Text,
//       function(result) {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           const selectedText = result.value;
//           if (selectedText && selectedText.trim()) {
//             document.getElementById('selectedText').value = selectedText;
//           }
//         }
//       }
//     );
//   }
// }

//Function to load and display email preview
function loadEmailPreview() {
  const item = Office.context.mailbox.item;
  const previewDiv = document.getElementById('previewContent');
  
  try {
    const subject = item.subject || 'No Subject';
    const from = item.from ? (item.from.displayName || item.from.emailAddress) : 'Unknown';
    const date = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : 'Unknown';
    
    previewDiv.innerHTML = `<strong>From:</strong> ${from}<br><strong>Subject:</strong> ${subject}<br><strong>Date:</strong> ${date}<br><br>Click the button below to forward this email.`;
  } catch (error) {
    previewDiv.textContent = 'Email details will be included when forwarded.';
  }
}

//Function to get email body
function getEmailBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error('Could not retrieve email body'));
        }
      }
    );
  });
}

//Function to send entire email via backend
async function sendEntireEmail() {
  const statusDiv = document.getElementById('status');
  const sendButton = document.getElementById('sendEntireBtn');
  
  try {
    //Getting recipient email
    const recipientEmail = getRecipientEmail();
    
    //Disable button and show loading
    sendButton.disabled = true;
    sendButton.textContent = 'Forwarding email';
    showStatus('Retrieving email content', 'info');
    
    const item = Office.context.mailbox.item;
    
    //Getting email details
    const subject = item.subject || 'No Subject';
    const fromName = item.from ? item.from.displayName || item.from.emailAddress : 'Unknown';
    const fromEmail = item.from ? item.from.emailAddress : '';
    const date = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : 'Unknown';
    const toList = item.to ? item.to.map(r => r.displayName || r.emailAddress).join(', ') : 'Unknown';
    
    //Getting email body
    showStatus('Retrieving email body...', 'info');
    const body = await getEmailBody();
    
    console.log('Retrieved email body:');
    console.log('- Length:', body.length);
    console.log('- Preview (first 100 chars):', body.substring(0, 100));
    console.log('- Has special characters:', /[^\x20-\x7E\n\r\t]/.test(body));
    
    //Preparing forwarded email content (avoid special characters that might break JSON)
    const forwardedContent = `========== FORWARDED EMAIL ==========

From: ${fromName}
Email: ${fromEmail}
To: ${toList}
Date: ${date}
Subject: ${subject}

---------- Email Body ----------

${body}

====================================`;
    
    //Sending to backend server
    showStatus('Sending email...', 'info');
    const response = await fetch(BACKEND_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        //Extra header to bypass ngrok browser warning
        'ngrok-skip-browser-warning': 'true'
      },
      body: JSON.stringify({
        text: forwardedContent,
        subject: `FWD: ${subject}`,
        recipient: recipientEmail
      })
    });
    
    const result = await response.json();
    
    console.log('Server response:', result);
    
    if (response.ok && result.success) {
      showStatus(`✓ Email forwarded successfully to ${recipientEmail}`, 'success');
    } else {
      console.error('Backend error:', result);
      console.error('Status code:', result.statusCode);
      console.error('Error details:', result.error);
      showStatus('Failed to forward email: ' + (result.message || 'Unknown error') + 
                 (result.statusCode ? ` (Status: ${result.statusCode})` : ''), 'error');
    }
    
  } catch (error) {
    console.error('Error forwarding email:', error);
    showStatus('Error: ' + error.message, 'error');
  } finally {
    //Re-enabling button
    sendButton.disabled = false;
    sendButton.textContent = 'Forward Entire Email';
  }
}

//Function to send selected text
async function sendSelectedText() {
  const textToSend = document.getElementById('selectedText').value.trim();
  const statusDiv = document.getElementById('status');
  const sendButton = document.getElementById('sendSelectedBtn');
  
  //Validating text
  if (!textToSend) {
    showStatus('Please paste or select some text first!', 'error');
    return;
  }
  
  try {
    //Get recipient email
    const recipientEmail = getRecipientEmail();
    
    //Disabling button and show loading
    sendButton.disabled = true;
    sendButton.textContent = 'Sending';
    showStatus('Sending selected text', 'info');
    
    const item = Office.context.mailbox.item;
    const subject = item.subject || 'No Subject';
    
    //Sending to backend server
    const response = await fetch(BACKEND_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        //Extra header to bypass ngrok browser warning
        'ngrok-skip-browser-warning': 'true'
      },
      body: JSON.stringify({
        text: textToSend,
        subject: `Selected text from: ${subject}`,
        recipient: recipientEmail
      })
    });
    
    const result = await response.json();
    
    if (response.ok && result.success) {
      showStatus(`✓ Selected text sent successfully to ${recipientEmail}`, 'success');
      //Clearing the text area after successful send
      setTimeout(() => {
        document.getElementById('selectedText').value = '';
      }, 2000);
    } else {
      console.error('Backend error:', result);
      showStatus('Failed to send: ' + (result.message || 'Unknown error'), 'error');
    }
    
  } catch (error) {
    console.error('Error sending text:', error);
    showStatus('Error: ' + error.message, 'error');
  } finally {
    //Re-enabling button
    sendButton.disabled = false;
    sendButton.textContent = 'Forward Selected Text';
  }
}

//Function to show status messages
function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = type;
  statusDiv.style.display = 'block';
  
  //Auto-hiding after 5 seconds for success messages
  if (type === 'success') {
    setTimeout(() => {
      statusDiv.style.display = 'none';
    }, 5000);
  }
}