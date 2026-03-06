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
            // ----- Replace the placeholders with real values -----
            string inputFilePath = "sample.docx";               // path to the source document
            string smtpHost = "smtp.example.com";               // SMTP server
            int smtpPort = 587;                                   // SMTP port (usually 587 or 465)
            string smtpUser = "user@example.com";               // SMTP user (optional)
            string smtpPassword = "password";                  // SMTP password (optional)
            string fromAddress = "sender@example.com";          // sender e‑mail
            string toAddress = "recipient@example.com";         // recipient e‑mail
            string subject = "Converted PDF Document";          // e‑mail subject
            string body = "Please find the PDF attached.";      // e‑mail body

            var sender = new DocumentEmailSender();
            sender.SendConvertedDocumentByEmail(
                inputFilePath,
                smtpHost,
                smtpPort,
                smtpUser,
                smtpPassword,
                fromAddress,
                toAddress,
                subject,
                body);
        }
    }

    public class DocumentEmailSender
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
            string subject,
            string body)
        {
            // Load the source document.
            Document doc = new Document(inputFilePath);

            // Convert to PDF and keep the result in a MemoryStream.
            using (MemoryStream pdfStream = new MemoryStream())
            {
                // Rule: use Save(stream, SaveFormat) overload.
                doc.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // reset for reading.

                // Prepare the e‑mail.
                using (MailMessage message = new MailMessage())
                {
                    message.From = new MailAddress(fromAddress);
                    message.To.Add(new MailAddress(toAddress));
                    message.Subject = subject;
                    message.Body = body;

                    string attachmentName = Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf";

                    // Attach the PDF – dispose the attachment with a using block.
                    using (Attachment pdfAttachment = new Attachment(pdfStream, attachmentName, "application/pdf"))
                    {
                        message.Attachments.Add(pdfAttachment);

                        // Send via SMTP.
                        using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                        {
                            if (!string.IsNullOrEmpty(smtpUser))
                            {
                                smtp.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                            }
                            else
                            {
                                smtp.UseDefaultCredentials = true;
                            }

                            // Most modern servers require SSL/TLS.
                            smtp.EnableSsl = true;
                            smtp.Send(message);
                        }
                    }
                }
            }
        }
    }
}
