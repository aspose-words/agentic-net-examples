using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Example
{
    public class DocumentEmailSender
    {
        /// <summary>
        /// Loads a document, converts it to DOC format and sends it as an email attachment.
        /// </summary>
        public void SendDocumentAsEmail(
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
            // Load the source document using Aspose.Words (lifecycle rule).
            Document doc = new Document(inputFilePath);

            // Convert the document to DOC format and write it into a memory stream.
            using (MemoryStream docStream = new MemoryStream())
            {
                doc.Save(docStream, SaveFormat.Doc);
                docStream.Position = 0; // Reset stream position for reading.

                // Prepare the e‑mail message.
                using (MailMessage message = new MailMessage())
                {
                    message.From = new MailAddress(fromAddress);
                    message.To.Add(new MailAddress(toAddress));
                    message.Subject = subject;
                    message.Body = body;

                    // Attach the converted DOC document.
                    string attachmentName = Path.GetFileNameWithoutExtension(inputFilePath) + ".doc";
                    Attachment attachment = new Attachment(docStream, attachmentName, "application/msword");
                    message.Attachments.Add(attachment);

                    // Configure the SMTP client.
                    using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
                    {
                        smtpClient.EnableSsl = true; // Enable SSL if required.
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

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – replace with real values or supply them via command‑line arguments.
            string inputFilePath = args.Length > 0 ? args[0] : "sample.docx";
            string smtpHost = args.Length > 1 ? args[1] : "smtp.example.com";
            int smtpPort = args.Length > 2 ? int.Parse(args[2]) : 587;
            string smtpUser = args.Length > 3 ? args[3] : null;
            string smtpPassword = args.Length > 4 ? args[4] : null;
            string fromAddress = args.Length > 5 ? args[5] : "sender@example.com";
            string toAddress = args.Length > 6 ? args[6] : "recipient@example.com";
            string subject = args.Length > 7 ? args[7] : "Document attached";
            string body = args.Length > 8 ? args[8] : "Please find the DOC document attached.";

            var sender = new DocumentEmailSender();
            sender.SendDocumentAsEmail(
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
}
