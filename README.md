# Outlook-desktop-custom-add-in-
Created a custom outlook add-in that helps forward entire mail or just portion of mail to desired email addresses.


## Pre-requisites

Node js
SendGrid API
Two ngrok API (two accounts)-> one for backend and one for front end


## Steps on running this add-in

Step 1 - download node js depending upon Operating System of your machine.

Step 2 - download ngrok and add it to environment variables-path (you can check whether ngrok is added
to your path by opening command terminal and typing ngrok --version, if it gives a value then success).

Step 3 - open a command terminal in the folder where manifest.xml, taskpane.js, etc., files are
present and run following two commands - "npx -y mkcert create-ca" then "npx -y mkcert create-cert".
Then right click on the cert.crt that has been created, select local machine option under Store Location;
Under Place all certificated in following store, browse and choose Trusted Root Certification 
Authorities.

Step 4 - open two command terminals in the same folder. On one terminal run the command 
"http-server -p 3000 --cors -S -C cert.crt -K cert.key" to start frontend and on the second terminal 
run the command "node server.js" to start backend. These terminals needs to be kept running.  

Step 5 - once all these are set up and running, open two more terminals anywhere, and run ngrok,
once on localhost:3000 and other on localhost:3001 simultaneously, using commands
"ngrok config add-authtoken your-authtoken" and then "ngrok http https://localhost:300x" 
where x = 0 or 1 based on localhost. 3001 is used for backend and 3000 is used for frontend.
The free domain looks like "https://crotched-kristine-overhighly.ngrok-free.dev" [frontend] or
"https://contrastable-overaffected-irma.ngrok-free.dev" [backend] and will be given as soon as ngrok
starts, just replace these domains in manifest.xml and taskpane.js accordingly.   

Step 6 - open outlook desktop, select add-in, select add custom add-in, then select manifest.xml,
icon will pop-up on the ribbon.

#Note- ngrok free tier provides usd 5 credits monthly, for production needs better hosting.
#Note- SendGrid API is only free for 60 days, if service is liked then you have to pay to use this API.
