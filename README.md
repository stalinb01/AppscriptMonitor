# AppscriptMonitor
Herramienta para la monitorización de servicios HTTP y HTTPS.

## Objetivo
Crear una herramienta capaz de monitorizar cualquier servicio web (HTTP/HTTPS) ubicado en cualquier lugar, utilizando Google Sheets y App Script.

## Recursos
### 1- Hoja de cálculo de Google con 4 hojas
### -Configuración: Esta hoja contiene los parámetros de configuración para el envío de mensajes por correo electrónico y Telegram. También incluye la lista de servidores con servicio HTTP/HTTPS a monitorizar.
    ### -Estado actual: Esta hoja muestra el estado del servidor (activo o inactivo) y la fecha y hora en que cambió de estado. 
    ### -Historial: Esta hoja contiene todos los registros de cambios de estado de todos los servidores. 
    ### -Informe de disponibilidad: Este es un resumen que muestra el tiempo que el servidor ha estado activo o inactivo.
### 2- Bot de Telegram: Con la herramienta Botfather de Telegram, crearemos un bot personalizado para enviar mensajes.
### 3- Cuentas de correo electrónico: Cuentas de correo electrónico para enviar mensajes cuando cambie el estado de los servidores.
### 4- App Script: Código con la lógica necesaria para obtener los datos y enviar los mensajes cuando cambie el estado del servidor.



=========================================================================================================================================

# AppscriptMonitor
Tool for Monitoring Http and Https Service.

## Objetive
Create a tool capable of monitoring any web service (HTTP/HTTPS) located anywhere, using Google Sheets and AppScript

## Resources

### 1- Google Sheet with 4 sheets
    ### -Configuration: This sheet contains the configuration parameters for sending messages via email and Telegram. It also includes the list of servers with HTTP/HTTPS service to be monitored
    ### -Current State: This sheet show server state, active or inactive and the date and time when the state changed. 
    ### -Story: This sheet contains all registers with state changes of all servers.
    ### -Availability report: This is a summary showing the time the server is active or inactive.
### 2- Telegram Bot: With the Botfather telegram tool, we will create a custom Bot to send messages through it.
### 3- Emails: Emails account to send messages when servers state changed
### 4- App Script: Code with the logic necessary to control get de data and send the messages when the server state changed.

