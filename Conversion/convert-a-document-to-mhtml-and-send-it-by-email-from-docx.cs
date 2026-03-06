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
        // Load the source DOCX file.
        string sourcePath = @"C:\Docs\input.docx";
        Document doc = new Document(sourcePath);

        // Convert the document to MHTML and keep it in memory.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Use the provided Save overload that accepts a stream and a SaveFormat.
            doc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream for reading.

            // Create the email message.
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("sender@example.com");
            mail.To.Add("recipient@example.com");
            mail.Subject = "Document in MHTML format";
            mail.Body = "Please find the attached MHTML document.";

            // Attach the MHTML content. The MIME type "message/rfc822" works for .mht files.
            Attachment attachment = new Attachment(mhtmlStream, "document.mht", "message/rfc822");
            mail.Attachments.Add(attachment);

            // Configure the SMTP client (replace with real server details).
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                smtp.EnableSsl = true;

                // Send the email.
                smtp.Send(mail);
            }
        }
    }
}
