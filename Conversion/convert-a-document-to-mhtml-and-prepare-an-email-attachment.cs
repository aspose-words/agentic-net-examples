using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlEmail
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string sourceDocPath = @"C:\Docs\SourceDocument.docx";

            // Path where the MHTML file will be saved.
            string mhtmlPath = @"C:\Docs\ConvertedDocument.mht";

            // Load the Word document using the provided Document constructor.
            Document doc = new Document(sourceDocPath);

            // Configure save options for MHTML output.
            // Use the HtmlSaveOptions constructor that accepts a SaveFormat.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources to improve compatibility with email clients.
                ExportCidUrlsForMhtmlResources = true,

                // Optional: make the output more readable.
                PrettyFormat = true
            };

            // Save the document as MHTML using the Document.Save method that accepts SaveOptions.
            doc.Save(mhtmlPath, saveOptions);

            // Prepare an email with the MHTML file as an attachment.
            using (MailMessage message = new MailMessage())
            {
                message.From = new MailAddress("sender@example.com");
                message.To.Add("recipient@example.com");
                message.Subject = "Converted MHTML Document";
                message.Body = "Please find the converted MHTML document attached.";

                // Attach the MHTML file.
                Attachment attachment = new Attachment(mhtmlPath);
                message.Attachments.Add(attachment);

                // Configure the SMTP client (adjust host, port, and credentials as needed).
                using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
                {
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential("username", "password");

                    // Send the email.
                    smtp.Send(message);
                }
            }

            Console.WriteLine("Document converted to MHTML and emailed successfully.");
        }
    }
}
