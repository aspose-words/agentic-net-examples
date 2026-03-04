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
        string mhtmlFilePath = @"C:\Input\sample.mht";

        // Load the MHTML document using Aspose.Words.
        Document doc = new Document(mhtmlFilePath);

        // Save the document to a memory stream in MHTML format.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create an attachment from the MHTML stream.
            Attachment attachment = new Attachment(
                mhtmlStream,
                "sample.mht",                     // File name for the attachment.
                "message/rfc822");                // MIME type for MHTML.

            // Build the mail message.
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "MHTML Document Attached",
                Body = "Please find the MHTML document attached."
            };
            mail.To.Add("recipient@example.com");
            mail.Attachments.Add(attachment);

            // Example: send the email (requires a valid SMTP server).
            // using (SmtpClient client = new SmtpClient("smtp.example.com"))
            // {
            //     client.Credentials = new System.Net.NetworkCredential("user", "password");
            //     client.Send(mail);
            // }

            // Dispose resources.
            attachment.Dispose();
            mail.Dispose();
        }
    }
}
