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
        // Load an existing document (any supported format).
        Document doc = new Document("InputDocument.docx");

        // Save the document to a memory stream in DOC format.
        using (MemoryStream docStream = new MemoryStream())
        {
            doc.Save(docStream, SaveFormat.Doc);
            docStream.Position = 0; // Reset stream for reading.

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Converted DOC Attachment";
            message.Body = "Please find the converted DOC file attached.";

            // Attach the DOC stream to the email.
            Attachment attachment = new Attachment(docStream, "ConvertedDocument.doc", "application/msword");
            message.Attachments.Add(attachment);

            // Configure and send the email via SMTP.
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                smtp.EnableSsl = true;
                smtp.Send(message);
            }
        }
    }
}
