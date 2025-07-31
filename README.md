# email_catcher
Basic program for catching emails that end up in my junk folder.

## Azure App Registration
1. Create app registration in Azure Portal
2. Set "Allow public client flows" to "Yes" in "Authentication" section.
    1. *May require changing page format to old to find this*
3. Add delegated permissions: Mail.Read, Mail.ReadWrite
4. Copy CLIENT_ID and TENANT_ID to .env file
