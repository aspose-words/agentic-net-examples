using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class EmailDocSender
{
    static void Main()
    {
        // Create a new Word document (using the create rule)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a test document.");

        // Save the document to a memory stream in DOC format (using the save rule)
        using (MemoryStream docStream = new MemoryStream())
        {
            DocSaveOptions saveOptions = new DocSaveOptions
            {
                SaveFormat = SaveFormat.Doc // Explicitly set DOC format
            };
            doc.Save(docStream, saveOptions);
            docStream.Position = 0; // Reset stream position for reading

            // Prepare the email message
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Test Email with DOC attachment";
            message.Body = "Please find the attached DOC file.";

            // Attach the DOC stream to the email
            Attachment attachment = new Attachment(docStream, "TestDocument.doc", "application/msword");
            message.Attachments.Add(attachment);

            // Configure the SMTP client (adjust host, port, credentials as needed)
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");

                // Send the email
                smtp.Send(message);
            }
        }
    }
}
