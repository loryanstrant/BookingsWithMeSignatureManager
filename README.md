# BookingsWithMeSignatureManager
A solution for end-users to easily add/remove their Bookings with Me public page to their email signature, using a Power App front-end, Power Automate workflows, Microsoft Graph calls, and an Exchange PowerShell script running in an Azure Automation Runbook

*** NOT COMPLETE***

While this solution does use extensionAttribute1, you can change this to use another one if this is already in use.

Prerequisites:
- Azure Automation account
  - Exchange Online Management module for PowerShell installed
  - Run As account created
  - Exchange Administrator role assigned<br>
  Follow this blog for steps on how to achieve the above steps: https://practical365.com/use-azure-automation-exchange-online/

Steps required to install:
1. Import the solution into Power Apps / Power Automate
2. Using the associated App Registration from your Azure Run As app in your Azure Active Directory
   - Add the "Exchange.ManageAsApp" application permission from Office 365 Exchange Online
   - Add the "User.ReadWrite.All" application permission Microsoft Graph
   - Grant admin consent for the permissions
   - Create a client secret
3. Create a Runbook in Azure Automation with the PowerShell script that utilises the Service Principal created in step 2
4. Update the two Azure Automation actions in the workflow "USERS - Book with Me (CHILD FLOW) - get ExchangeGUID" to use your new Runbook (expect errors, they will go away as you update the fields)
5. Update the corresponding variables in the following workflows with your Tenant ID, Client / Application ID, and Secret obtained in step 2:
   - "USERS - Book with Me - clear extensionAttribute1"
   - "USERS - Book with Me - get status of extensionAttribute1"
   - "USERS - Book with Me - set URL in extensionAttribute1"
6. Turn on the workflow "USERS - Book with Me (CHILD FLOW) - get ExchangeGUID" and copy the HTTP POST URL from the trigger
7. Paste the HTTP POST URL from step 6 into the "HTTP - call child workflow" action in the workflow "USERS - Book with Me - set URL in extensionAttribute1"
