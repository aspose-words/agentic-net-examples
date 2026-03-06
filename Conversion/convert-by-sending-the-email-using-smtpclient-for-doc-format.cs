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

        // Add some content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Save the document to a memory stream in DOC format.
        using (MemoryStream docStream = new MemoryStream())
        {
            doc.Save(docStream, SaveFormat.Doc);   // Use the provided Save(Stream, SaveFormat) method.
            docStream.Position = 0;                // Reset stream position for reading.

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Document attached";
            message.Body = "Please find the DOC attached.";

            // Attach the DOC stream to the email.
            Attachment attachment = new Attachment(docStream, "Document.doc", "application/msword");
            message.Attachments.Add(attachment);

            // Configure the SMTP client.
            SmtpClient smtp = new SmtpClient("smtp.example.com", 587)
            {
                Credentials = new NetworkCredential("username", "password"),
                EnableSsl = true
            };

            // Send the email.
            smtp.Send(message);
        }
    }
}
