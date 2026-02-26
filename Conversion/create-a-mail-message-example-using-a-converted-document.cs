using System;
using Aspose.Words;
using Aspose.Words.Settings;

class MailMessageExample
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple mail merge template.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Configure mail‑merge settings so that the result will be an e‑mail.
        MailMergeSettings settings = doc.MailMergeSettings;
        settings.Destination = MailMergeDestination.Email;          // Generate e‑mail.
        settings.MailSubject = "Welcome to Aspose.Words";           // Subject line.
        settings.AddressFieldName = "EmailAddress";                 // Column that contains e‑mail addresses.
        settings.MailAsAttachment = false;                         // Put the merged document in the e‑mail body.

        // Normally you would also set the data source (e.g., a CSV file) and execute the merge.
        // For this example we only demonstrate the configuration and conversion.

        // Save the document. When opened in Microsoft Word the mail merge will produce an e‑mail.
        doc.Save("MailMessageExample.docx");
    }
}
