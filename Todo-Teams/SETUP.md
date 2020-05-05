# Setup

## ngrok
- Create a named instance using ngrok.io (alternatively use a random ngrok adress if you don't want to pay a few dollars)
- update `.env` and `HOSTNAME` with your ngrok address

## create bot
- Go to https://portal.azure.com
- Create a new Resource Group
- Add a `Bot Channels Registration` to the Resource Group
- Configure the bot and use `https://<ngrokaddress>.ngrok.io/api/messages` as the Messaging endpoint
- Click on Create
- When the bot is created, go to the bot in the Azure Portal and then to Settings
- Copy the Microsoft App Id and add that to the `MICROSOFT_APP_ID` setting in the `.env` file
- Go to *Azure Active Directory* and *App Registrations (Preview)* in the Azure Portal
- Locate the bot you just created and select it, then select *Certificates and Secrets*
- Create a new *Client Secret*
- Copy the generated secret and add that to the `MICROSOFT_APP_PASSWORD` setting in the `.env` file
- Go to *API Permissions* and choose *Microsoft Graph* > *Delegated Permissions* and add `Tasks.ReadWrite`, `User.Read`, `openid` and `profile`
- Under *Authentication* as redirect add `https://token.botframework.com/.auth/web/redirect` for the type *Web*
- Under *Implicit grant*, check both *Access tokens* and *ID tokens*
- Click save
- Go back to the bot and click on Channels
- Click on the Microsoft Teams icon to add the Microsoft Teams channel and click Save (and agree to the terms of service)
- Click on Settings
- Click on *Add Setting* under *OAuth Connection Settings*
- Add the name `AADv2` and choose `Azure Active Directory v2` as the provider
- As *Client Id* use the `MICROSOFT_APP_ID` from the `.env` file and for *Client secret* use the `MICROSOFT_APP_PASSWORD` from the `.env` file (or use a new custom client secret). For *Tenand ID* specify `common` or your custom domain name and for *Scopes* specify `openid, profile, Tasks.ReadWrite, User.Read`

## connector
 - TODO



