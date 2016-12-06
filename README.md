# OutlookAddin

I deploy this by:

1. Install Azure CLI: https://docs.microsoft.com/en-us/azure/xplat-cli-install
2. azure config mode asm
3. azure login
4. azure site create -s <subscription id> --git {appname}
5. git push azure master

On the Azure Portal (https://portal.azure.com):

1. Select 'App Services' under 'Web + Mobile'
2. Select the app named '{appname}'
3. Under Settings->Application settings update the 'Virtual applications and directories' so '/' points to 'site\wwwroot\appread'

I'm sure there is a better way, but this was the quick and easy way I got this working.

Don't forget that your app in Azure AD will need the delegated permission for 'Read all groups' on the Microsoft Graph application under 'Permissions to other applications'.
