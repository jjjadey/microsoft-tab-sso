# About project: Tab app with microsoft graph api

This project is tabs app add-in for Microsoft Teams.
* Client side: ReactJS
* Server side: NodeJS

## Prerequisites

* [NodeJS](https://nodejs.org/en/)
* An M365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
* [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version after 1.55 or [TeamsFx CLI](https://aka.ms/teamsfx-cli)

# How to run in localhost

If you want to run in localhost, you can use 2 options.

1. With server-side code
2. Without server-side code: SSO authentication provided by teams toolkit

## 1\. With server\-side code

1. Install package `npm run installAll`
2. Register new App in [Portal Azure](https://portal.azure.com/)
You can hit the `F5` key in VS Code (Teams tookit will auto create new app and setup everything) or create by yourself [Register an Azure AD App](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
3. Edit `.env` for tabs

* Open folder `tabs`
* Open your app in 2. and copy `Application (client) ID`
* copy `.env.sample` and rename file to `.env`. Then paste your App id in `REACT_APP_CLIENT_ID`
![appid](https://user-images.githubusercontent.com/47839144/150120152-4c2c4ba6-48f5-424c-a946-4261b1742a6a.jpg)

4. Upadate `Application (client) ID` in `auth-start.html` (client side) 
5. Edit `.env` for server

* Open folder `server`
* Open your app in 2. and copy `Certificates value` (3rd column in below figure). If you can't copy value, you can create new one (click `New client secret`)
* copy `.env.sample` and rename file to `.env`. Then paste your App id in `CLIENT_ID` and Certificates value in `CLIENT_SECRET`
![cer](https://user-images.githubusercontent.com/47839144/150120677-ff5e5dde-d39b-4c81-b079-3d9bd67a0ae0.jpg)

6. Run client side (tabs) in local
`cd tabs` and `npm run start`
7. Run server side in local
`cd server` and `npm run start`
8. Upload Manifest
    * go to folder `manifest`
    * copy file `manifest.sample.json` and rename to `manifeest.json`
    * cd .fx/config/localSetting.json and copy value
    `{{localSettings.teamsApp.teamsAppId}}`= teamsAppId
    `{{localSettings.auth.clientId}}` = clientId
    `{{{localSettings.auth.applicationIdUris}}}` = applicationIdUris
    * If you have no localSetting.json, hitting `F5` or you can copy value from your app overview in Portal Azure (teamsAppId can be any value in format xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)
    * Compress manifest.json and folder resources to zip file.
    * open MS Teams app in browser. Then click Apps> Upload a custom app and upload your zip file
  
![upload](https://user-images.githubusercontent.com/47839144/150121038-7d76e353-96ae-426a-ae84-545744941860.jpg)
![app](https://user-images.githubusercontent.com/47839144/150121289-5e4d6c9e-d33c-4805-9fc7-b262265640eb.jpg)

## 2\. Without server\-side code

* npm run installAll
* hitting the `F5` key in Visual Studio Code.

# How to run server side with Ngrok
1. `cd server` and `npm run start`
2. Install ngrok and run `ngrok http 55000`
![ngrok](https://user-images.githubusercontent.com/47839144/150123726-c41e2eb2-d50a-4dc2-b071-ddc7c25f8d7e.jpg)
3. Copy your path forwarding `https://xxxx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx-xxxx.ngrok.io`

# How to host client side on Firebase
[Firebase Tutorial - How to Set Up and Deploy with ReactJS](https://www.youtube.com/watch?v=1wZw7RvXPRU)
1. `cd tabs` 
2. Copy `.env.sample.production` and rename to `.env.production`
3. Change `REACT_APP_TEAMSFX_ENDPOINT` to ngrok path forwarding
4. `npm run build:production`
5. `firebase deploy`

NOTE: If you use free version ngrok, the path was change after you terminate the server side. 
So you must update `.env.production`, `npm run build:production`, and deploy


# How to test on production
1. `cd manifestProd`
2. copy file `manifest.sample.json` and rename to `manifest.json`
3. edit `https://tab-sso-client.web.app` to your client domain
4. edit `webApplicationInfo.resource` to  `api://{your-client-domain}/{your-app-client-id-from-portalAzure}` 
such as `api://tab-sso-client.web.app/84a71487-45a8-4504-8aa5-e5xxxxxxxxxx`
5. Add `redirectUri` in Portal Azure
![redirectUri](https://user-images.githubusercontent.com/47839144/150128024-c278e6fc-25da-45b2-85d7-245be227c1b3.jpg)
NOTE: It must be your client domain

6. Update `Expose API` in portal Azure. The value be the same as 4.
![exposeApi](https://user-images.githubusercontent.com/47839144/150125884-32c5f8a4-fdd1-4e1a-a340-a45f7020dd72.jpg)
NOTE: If `webApplicationInfo.resource` not be the same as `expose api` in azure portal, the error resourceDisabled was found.
7. Compress manifefst to zip and upload in MS Teams for testing.


## Edit the manifest

You can find the Teams app manifest in `templates/appPackage` folder. The folder contains two manifest files:

* `manifest.local.template.json`: Manifest file for Teams app running locally.
* `manifest.remote.template.json`: Manifest file for Teams app running remotely (After deployed to Azure).

Both files contain template arguments with `{...}` statements which will be replaced at build time. You may add any extra properties or permissions you require to this file. See the [schema reference](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema) for more information.

## Deploy to Azure

Deploy your project to Azure by following these steps:

| From Visual Studio Code | From TeamsFx CLI |
| ----------------------- | ---------------- |
| <ul><li>Open Teams Toolkit, and sign into Azure by clicking the `Sign in to Azure` under the `ACCOUNTS` section from sidebar.</li><li>After you signed in, select a subscription under your account.</li><li>Open the Teams Toolkit and click `Provision in the cloud` from DEVELOPMENT section or open the command palette and select: `Teams: Provision in the cloud`.</li><li>Open the Teams Toolkit and click `Deploy to the cloud` or open the command palette and select: `Teams: Deploy to the cloud`.</li></ul> | <ul><li>Run command `teamsfx account login azure`.</li><li>Run command `teamsfx account set --subscription <your-subscription-id>`.</your-subscription-id></li><li>Run command `teamsfx provision`.</li><li>Run command: `teamsfx deploy`.</li></ul> |

> Note: Provisioning and deployment may incur charges to your Azure Subscription.

## Validate manifest file

To check that your manifest file is valid:

* From Visual Studio Code: open the Teams Toolkit and click `Validate manifest file` or open the command palette and select: `Teams: Validate manifest file`.
* From TeamsFx CLI: run command `teamsfx validate` in your project directory.
