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
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a DOCM document sent via email.");

        // Save the document to a memory stream in DOCM format.
        using (MemoryStream docStream = new MemoryStream())
        {
            doc.Save(docStream, SaveFormat.Docm);
            docStream.Position = 0; // Reset stream position for reading.

            // Build the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Aspose.Words DOCM Email";
            message.Body = "Please find the attached DOCM document.";

            // Attach the DOCM document from the memory stream.
            Attachment attachment = new Attachment(
                docStream,
                "Document.docm",
                "application/vnd.ms-word.document.macroEnabled.12");
            message.Attachments.Add(attachment);

            // Configure and send the email using SmtpClient.
            using (SmtpClient client = new SmtpClient("smtp.example.com", 587))
            {
                client.Credentials = new NetworkCredential("username", "password");
                client.EnableSsl = true;
                client.Send(message);
            }
        }
    }
}
