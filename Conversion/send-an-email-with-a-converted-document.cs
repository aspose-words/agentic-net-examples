using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SendConvertedDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            // TODO: replace the following placeholders with real values or read them from configuration.
            string inputFilePath = "sample.docx";               // path to the source Word document
            string smtpHost = "smtp.example.com";               // SMTP server
            int smtpPort = 587;                                   // SMTP port (usually 25, 465 or 587)
            string smtpUser = "user@example.com";               // SMTP username (optional)
            string smtpPassword = "password";                   // SMTP password (optional)
            string fromAddress = "sender@example.com";          // sender e‑mail address
            string toAddress = "recipient@example.com";         // recipient e‑mail address
            string subject = "Converted PDF Document";          // e‑mail subject

            var sender = new EmailSender();
            sender.SendConvertedDocumentByEmail(
                inputFilePath,
                smtpHost,
                smtpPort,
                smtpUser,
                smtpPassword,
                fromAddress,
                toAddress,
                subject);
        }
    }

    public class EmailSender
    {
        /// <summary>
        /// Loads a Word document, converts it to PDF in memory, and sends it as an e‑mail attachment.
        /// </summary>
        public void SendConvertedDocumentByEmail(
            string inputFilePath,
            string smtpHost,
            int smtpPort,
            string smtpUser,
            string smtpPassword,
            string fromAddress,
            string toAddress,
            string subject)
        {
            // Load the source document.
            Document doc = new Document(inputFilePath);

            // Convert to PDF and keep the bytes in a MemoryStream.
            using (MemoryStream pdfStream = new MemoryStream())
            {
                doc.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // rewind for reading

                // Prepare e‑mail objects inside using blocks to guarantee disposal.
                using (MailMessage message = new MailMessage())
                using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
                {
                    message.From = new MailAddress(fromAddress);
                    message.To.Add(new MailAddress(toAddress));
                    message.Subject = subject;
                    message.Body = "Please find the converted PDF document attached.";

                    string attachmentName = Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf";
                    Attachment pdfAttachment = new Attachment(pdfStream, attachmentName, "application/pdf");
                    message.Attachments.Add(pdfAttachment);

                    smtpClient.EnableSsl = true; // adjust if the server does not require SSL
                    if (!string.IsNullOrEmpty(smtpUser))
                    {
                        smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                    }

                    smtpClient.Send(message);
                }
            }
        }
    }
}
