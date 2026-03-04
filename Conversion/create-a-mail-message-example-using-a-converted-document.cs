using System;
using Aspose.Words;
using Aspose.Words.Settings;

class MailMergeEmailExample
{
    static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Build the document content with merge fields.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Configure mail merge settings so that Word will generate an e‑mail
        // when the document is opened.
        MailMergeSettings settings = doc.MailMergeSettings;
        settings.Destination = MailMergeDestination.Email;          // Generate e‑mail
        settings.MailSubject = "Hello ${FirstName}";                // Subject line (Word will replace the placeholder)
        settings.MailAsAttachment = true;                          // Send the merged document as an attachment
        settings.AddressFieldName = "Email";                        // Column that contains the recipient address

        // (Optional) Specify the main document type as e‑mail.
        settings.MainDocumentType = MailMergeMainDocumentType.Email;

        // Save the document. When opened in Microsoft Word, it will prompt to send an e‑mail.
        doc.Save("MailMergeEmailExample.docx");
    }
}
