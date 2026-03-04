using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example values – replace with real data or pass via command‑line arguments.
            string docxPath = "Sample.docx";
            string smtpHost = "smtp.example.com";
            int smtpPort = 587;
            string smtpUser = "user@example.com";
            string smtpPassword = "password";
            string fromAddress = "sender@example.com";
            string toAddress = "recipient@example.com";
            string subject = "Document as MHTML";
            string body = "Please find the document attached as MHTML.";

            var sender = new MhtmlEmailSender();
            sender.SendDocxAsMhtml(docxPath, smtpHost, smtpPort, smtpUser, smtpPassword,
                                   fromAddress, toAddress, subject, body);
        }
    }

    public class MhtmlEmailSender
    {
        /// <summary>
        /// Loads a DOCX file, converts it to MHTML, and sends it as an e‑mail attachment.
        /// </summary>
        public void SendDocxAsMhtml(
            string docxPath,
            string smtpHost,
            int smtpPort,
            string smtpUser,
            string smtpPassword,
            string fromAddress,
            string toAddress,
            string subject,
            string body)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Save the document to a memory stream in MHTML format.
            using (MemoryStream mhtmlStream = new MemoryStream())
            {
                doc.Save(mhtmlStream, SaveFormat.Mhtml);
                mhtmlStream.Position = 0; // Reset for reading.

                // Build the e‑mail message.
                using (MailMessage message = new MailMessage())
                {
                    message.From = new MailAddress(fromAddress);
                    message.To.Add(toAddress);
                    message.Subject = subject;
                    message.Body = body;
                    message.IsBodyHtml = false;

                    // Attach the MHTML content. "text/html" is the appropriate MIME type for MHTML.
                    Attachment attachment = new Attachment(mhtmlStream, "Document.mhtml", "text/html");
                    message.Attachments.Add(attachment);

                    // Configure and send via SMTP.
                    using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                    {
                        smtp.EnableSsl = true;
                        smtp.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                        smtp.Send(message);
                    }
                }
            }
        }
    }
}
