# Read Office 365 Mailbox and Download Attachments using C# and Microsoft.Exchange.WebServices

This repository contains a sample C# application that demonstrates how to use the Microsoft Exchange Web Services (EWS) Managed API to read emails from an Office 365 mailbox, download email attachments, and move the processed emails to an archive folder.

## Features

- Connect to an Office 365 mailbox using Exchange Web Services (EWS).
- Retrieve emails from the inbox.
- Download email attachments.
- Move processed emails to an archive folder.

## Prerequisites

- Office 365 Account: You will need an Office 365 mailbox with necessary permissions.
- Microsoft.Exchange.WebServices: The project uses the EWS Managed API to interact with the Office 365 mailbox. Ensure the EWS Managed API is installed via NuGet.
- C#: Basic knowledge of C# programming and .NET is required to understand the code.

## Getting Started

- Clone the Repository
```
git clone [https://github.com/your-username/your-repository.git
cd your-repository](https://github.com/raselahmedit09/ReadOffice365Mailbox.git)
```
- Install Dependencies
```
Open the solution in Visual Studio and install the Microsoft.Exchange.WebServices NuGet package if not already installed.

> Install-Package Microsoft.Exchange.WebServices
```
- Configuration
```
Update the following configuration in the application to match your Office 365 mailbox settings:

Exchange URL: The URL to your Exchange server (e.g., https://outlook.office365.com/EWS/Exchange.asmx).
Credentials: Set your Office 365 username and password.
Example of setting credentials in code:

var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
{
    Credentials = new WebCredentials("your-email@domain.com", "your-password"),
    Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
};
```

## Run the Application

Once the configuration is set, you can run the application to:

Read emails from the Office 365 inbox.
Download any attachments.
Move the processed emails to the designated archive folder.

### Sample Code Snippet
```
// Create Exchange Service instance
ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
service.Credentials = new WebCredentials("your-email@domain.com", "your-password");
service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

// Find emails in the inbox
FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(10));

// Loop through emails and download attachments
foreach (EmailMessage email in findResults.Items)
{
    email.Load();  // Load email details
    foreach (Attachment attachment in email.Attachments)
    {
        if (attachment is FileAttachment fileAttachment)
        {
            fileAttachment.Load("path-to-save/" + fileAttachment.Name);
        }
    }

    // Move email to archive folder after processing
    email.Move(WellKnownFolderName.ArchiveMsgFolderRoot);
}

```

## How it Works

- Connect to Office 365: The application uses the ExchangeService class from the EWS Managed API to connect to the Office 365 mailbox using the provided credentials.
- Retrieve Emails: The FindItems method retrieves a list of emails from the inbox.
- Download Attachments: If an email contains attachments, the attachments are downloaded to the specified folder.
- Move Emails: After processing the email and downloading attachments, the email is moved to the archive folder using the Move method.

## Additional Notes

- The EWS Managed API requires proper authentication, and your Office 365 account must have the necessary permissions to access mailboxes.
- Ensure that you handle credentials securely and avoid hardcoding sensitive information in the code.


