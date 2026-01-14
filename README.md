# AppscriptMonitor
Tools for Monitoring Http and Https Service.

## Objetive
Create a tool capable of monitoring any web service (HTTP/HTTPS) located anywhere, using Google Sheets and AppScript

## Resources

1- Google Sheet with 4 sheets
    - Configuration: This sheet contains the configuration parameters for sending messages via email and Telegram. It also includes the list of servers with HTTP/HTTPS service to be monitored
    - Current State: This sheet show server state, active or inactive and the date and time when the state changed. 
    - Story: This sheet contains all registers with state changes of all servers.
    - Availability report: This is a summary showing the time the server is active or inactive.

2- Telegram Bot: With the Botfather telegram tool, we will create a custom Bot to send messages through it.
3- Emails: Emails account to send messages when servers state changed
4- App Script: Code with the logic necessary to control get de data and send the messages when the server state changed.

