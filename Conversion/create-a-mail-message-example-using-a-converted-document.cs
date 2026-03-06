using System;
using Aspose.Words;
using Aspose.Words.Settings;

class MailMessageExample
{
    static void Main()
    {
        // Load an existing Word document that will be used as the mail‑merge template.
        // The Document constructor handles the creation and loading lifecycle.
        Document doc = new Document("Template.docx");

        // Access the mail‑merge settings of the document.
        MailMergeSettings settings = doc.MailMergeSettings;

        // Configure the document to be sent as an e‑mail after the mail merge.
        settings.Destination = MailMergeDestination.Email;          // Email output
        settings.MailAsAttachment = true;                           // Send the merged document as an attachment
        settings.MailSubject = "Your personalized document";        // Subject line of the e‑mail
        settings.AddressFieldName = "RecipientEmail";               // Column name that contains e‑mail addresses
        settings.MainDocumentType = MailMergeMainDocumentType.Email; // Mark the document as an e‑mail type

        // (Optional) If you have merge data, execute the mail merge here.
        // string[] fieldNames = { "FirstName", "LastName", "RecipientEmail" };
        // object[] fieldValues = { "John", "Doe", "john.doe@example.com" };
        // doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the configured document. The Save method determines the format from the file extension.
        doc.Save("ConvertedMailMessage.docx");
    }
}
