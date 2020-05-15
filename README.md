# MailChamp
An Email Client to Send and Receive Emails from your Office 365 Account.<br/>
***This application never stores user's credentials.***

Mail champ solution consists of 3 Visual Studio Projects. 
1) Class Library 
2) Console Application
3) WPF Application.

## EmailHelper - Class Library

This is a reusable library built using .Net Framework 4.7. 
1) EmailSender - A Class to Send the email by setting Email Object
2) EmailReader - A Class to read all the email as pet the Email Object Properties. Currently supports Inbox only.
3) ExchangeServiceUtil - A Class which connects to EWS API with a given token.

## MailChamp - Console Application

This is a sample console application to demonstrate the capability of the Class Library. 

This application has been designed to enable login to Microsoft account using Oauth and using explicit credentials. 
The Program asks user to login to his Office 365 account (a popup window opens), User enters his credentials and 
only token is passed back to our application from Microsoft.  
Same Token is used to fetch Mails from User's Inbox or to Send Emails via EWS.  
If the user decides to cancel the Office login window, application falls back to explicit credentials model. 

Application now asks the user to enter email and password in the console. Credentials are never stored on the server. They are only used for transacting with EWS and is automatically destroyed once the program ends.

## MailChampWin - A WPF Application

This is also a sample application to demonstrate the reusability of Class library. 

This application uses WPF.Net to show Inbox data on a GUI. This application only supports login via Oauth and also has a feature 
to cache the retrieved token and fetch a new token silently when the old token expires.


# *Build and Run*

This can be run only on Windows machine as this was built using .Net Framework 4.7. Machine must have this version of .Net installed.

Prior to opening the solution, Make sure you have set the ClientId and TenantId in your Environment. 
Get the ClientID from your Office 365 account (You must be the administrator for your org account). 

If you are a developer, Get a developer account from - https://developer.microsoft.com/en-us/microsoft-365/dev-program .

Register the app in the Azure portal using the same developer account (Use org admin account) to get the Client Id and Tenant ID.
More Info: https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications 

#### *ENV Variables to set*
AZ_ClientID: XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  
AZ_TenantId: XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

```
Open the MailChamp.sln solution file in Visual Studio 2019.
Build the application.
Nugets are restored automatically during the build.
You can set either the Console App or the WPF as a startup project.
Press F5 to Run the application
```
