using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Save the document to a memory stream in DOC format.
        using (MemoryStream docStream = new MemoryStream())
        {
            doc.Save(docStream, SaveFormat.Doc);
            docStream.Position = 0; // Reset stream for reading.

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Document attached";
            message.Body = "Please find the attached DOC file.";

            // Attach the document stream.
            Attachment attachment = new Attachment(docStream, "Document.doc", "application/msword");
            message.Attachments.Add(attachment);

            // Configure and send via SMTP.
            using (SmtpClient client = new SmtpClient("smtp.example.com", 587))
            {
                client.Credentials = new NetworkCredential("username", "password");
                client.EnableSsl = true;
                client.Send(message);
            }
        }
    }
}
