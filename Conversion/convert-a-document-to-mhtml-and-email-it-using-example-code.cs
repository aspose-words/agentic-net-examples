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
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Configure save options to produce MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true, // Use CID URLs for embedded resources.
            PrettyFormat = true                     // Make the output more readable.
        };

        // Save the document to a memory stream in MHTML format.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, saveOptions);
            mhtmlStream.Position = 0; // Reset stream for reading.

            // Create an email message.
            MailMessage message = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Document as MHTML",
                Body = "Please find the document attached as MHTML."
            };
            message.To.Add("recipient@example.com");

            // Attach the MHTML content.
            Attachment attachment = new Attachment(mhtmlStream, "Document.mht", "multipart/related");
            message.Attachments.Add(attachment);

            // Send the email via SMTP.
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new NetworkCredential("username", "password");
                smtp.EnableSsl = true;
                smtp.Send(message);
            }
        }
    }
}
