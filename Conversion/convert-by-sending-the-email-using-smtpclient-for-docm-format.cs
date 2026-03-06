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
        builder.Writeln("Hello, this is a DOCM attachment.");

        // Save the document to a memory stream in DOCM format.
        using (MemoryStream ms = new MemoryStream())
        {
            doc.Save(ms, SaveFormat.Docm);
            ms.Position = 0; // Reset stream position for reading.

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "DOCM Document";
            message.Body = "Please find the DOCM document attached.";

            // Attach the DOCM document from the memory stream.
            Attachment attachment = new Attachment(
                ms,
                "Document.docm",
                "application/vnd.ms-word.document.macroEnabled.12");
            message.Attachments.Add(attachment);

            // Configure the SMTP client.
            SmtpClient client = new SmtpClient("smtp.example.com", 587)
            {
                Credentials = new NetworkCredential("username", "password"),
                EnableSsl = true
            };

            // Send the email.
            client.Send(message);
        }
    }
}
