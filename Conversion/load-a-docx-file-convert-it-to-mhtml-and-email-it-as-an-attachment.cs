using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;

namespace AsposeWordsMhtmlEmail
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string sourceDocxPath = @"C:\Docs\SourceDocument.docx";

            // Email configuration (replace with real values).
            const string smtpHost = "smtp.example.com";
            const string smtpUser = "smtp_user";
            const string smtpPassword = "smtp_password";
            const string fromAddress = "sender@example.com";
            const string toAddress = "recipient@example.com";

            // Load the DOCX document using Aspose.Words.
            Document doc = new Document(sourceDocxPath);

            // Save the document to a memory stream in MHTML format.
            using (MemoryStream mhtmlStream = new MemoryStream())
            {
                doc.Save(mhtmlStream, SaveFormat.Mhtml);
                mhtmlStream.Position = 0; // Reset stream position for reading.

                // Prepare the e‑mail message.
                using (MailMessage message = new MailMessage())
                {
                    message.From = new MailAddress(fromAddress);
                    message.To.Add(toAddress);
                    message.Subject = "Converted MHTML Document";
                    message.Body = "Please find the attached MHTML file.";

                    // Attach the MHTML content. Use a generic MIME type.
                    Attachment attachment = new Attachment(mhtmlStream, "Document.mhtml", "application/octet-stream");
                    message.Attachments.Add(attachment);

                    // Send the e‑mail via SMTP.
                    using (SmtpClient smtpClient = new SmtpClient(smtpHost))
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(smtpUser, smtpPassword);
                        smtpClient.Send(message);
                    }
                }
            }
        }
    }
}
