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
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SendConvertedDocument <inputFilePath> <recipientEmail>");
                return;
            }

            string inputFilePath = args[0];
            string recipientEmail = args[1];

            try
            {
                DocumentEmailSender sender = new DocumentEmailSender();
                sender.SendConvertedDocumentByEmail(inputFilePath, recipientEmail);
                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class DocumentEmailSender
    {
        /// <summary>
        /// Loads a Word document, converts it to PDF, and sends it as an email attachment.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source .docx (or any supported) file.</param>
        /// <param name="recipientEmail">Email address of the recipient.</param>
        public void SendConvertedDocumentByEmail(string inputFilePath, string recipientEmail)
        {
            // Load the source document.
            Document doc = new Document(inputFilePath);

            // Convert the document to PDF in a memory stream.
            using (MemoryStream pdfStream = new MemoryStream())
            {
                doc.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0;

                // Prepare the email.
                using (MailMessage message = new MailMessage())
                {
                    message.From = new MailAddress("sender@example.com"); // TODO: replace with a valid sender address.
                    message.To.Add(recipientEmail);
                    message.Subject = "Your converted PDF document";
                    message.Body = "Please find the converted PDF attached.";

                    string attachmentName = Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf";

                    // Attach the PDF.
                    using (Attachment pdfAttachment = new Attachment(pdfStream, attachmentName, "application/pdf"))
                    {
                        message.Attachments.Add(pdfAttachment);

                        // Configure SMTP client.
                        using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                        {
                            smtp.EnableSsl = true;
                            smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                            smtp.Send(message);
                        }
                    }
                }
            }
        }
    }
}
