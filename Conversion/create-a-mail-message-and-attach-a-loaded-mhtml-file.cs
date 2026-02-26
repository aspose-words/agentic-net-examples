using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source MHTML file.
        string mhtmlPath = @"C:\Temp\source.mht";

        // Load the MHTML document using Aspose.Words.
        Document doc = new Document(mhtmlPath); // load

        // Save the document to a memory stream in MHTML format.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, SaveFormat.Mhtml); // save
            mhtmlStream.Position = 0; // reset for reading

            // Create a mail message.
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Attached MHTML Document",
                Body = "Please find the MHTML document attached."
            };
            mail.To.Add("recipient@example.com");

            // Attach the MHTML content. Use the appropriate MIME type.
            Attachment attachment = new Attachment(mhtmlStream, "document.mht", "message/rfc822");
            mail.Attachments.Add(attachment);

            // Send the e‑mail (SMTP settings must be configured for your environment).
            using (SmtpClient client = new SmtpClient("smtp.example.com"))
            {
                client.Send(mail);
            }
        }
    }
}
