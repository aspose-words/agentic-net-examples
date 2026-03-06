using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing MHTML file into an Aspose.Words Document.
        Document mhtmlDoc = new Document("input.mht");

        // Save the document back to MHTML format into a memory stream.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Use the SaveFormat.Mhtml to keep the original format.
            mhtmlDoc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream for reading.

            // Create a new e‑mail message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Document attached as MHTML";

            // Attach the MHTML content to the e‑mail.
            Attachment attachment = new Attachment(mhtmlStream, "document.mht", "message/rfc822");
            message.Attachments.Add(attachment);

            // Optional: send the e‑mail using an SMTP server.
            // using (SmtpClient client = new SmtpClient("smtp.example.com"))
            // {
            //     client.Credentials = new System.Net.NetworkCredential("user", "password");
            //     client.Send(message);
            // }
        }
    }
}
