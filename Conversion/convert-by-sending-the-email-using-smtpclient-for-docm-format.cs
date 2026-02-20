using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a DOCM document created with Aspose.Words.");

        // Configure save options for DOCM format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Save the document to a memory stream.
        using (MemoryStream docStream = new MemoryStream())
        {
            doc.Save(docStream, saveOptions);
            docStream.Position = 0; // Reset stream position for reading.

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "DOCM Document Attachment";
            message.Body = "Please find the attached DOCM document.";

            // Attach the DOCM document from the memory stream.
            Attachment attachment = new Attachment(
                docStream,
                "Document.docm",
                "application/vnd.ms-word.document.macroEnabled.12");
            message.Attachments.Add(attachment);

            // Configure the SMTP client.
            SmtpClient smtpClient = new SmtpClient("smtp.example.com", 587)
            {
                EnableSsl = true,
                Credentials = new NetworkCredential("smtp_user", "smtp_password")
            };

            // Send the email.
            smtpClient.Send(message);
        }
    }
}
