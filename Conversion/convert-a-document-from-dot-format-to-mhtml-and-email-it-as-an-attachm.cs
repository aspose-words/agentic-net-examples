using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToMhtmlEmail
{
    static void Main()
    {
        // Path to the source .dot (Word template) file.
        string sourceDotPath = @"C:\Input\Template.dot";

        // Load the .dot document. The Document constructor handles the creation and loading lifecycle.
        Document doc = new Document(sourceDotPath);

        // Convert the document to MHTML and store it in a memory stream.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Save using the SaveFormat enumeration for MHTML.
            doc.Save(mhtmlStream, SaveFormat.Mhtml);

            // Reset the stream position so it can be read from the beginning.
            mhtmlStream.Position = 0;

            // Prepare the email message.
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Converted MHTML Document";
            message.Body = "Please find the converted MHTML document attached.";

            // Attach the MHTML stream. The name given to the attachment will be the filename seen by the recipient.
            Attachment attachment = new Attachment(mhtmlStream, "Template.mhtml", "application/octet-stream");
            message.Attachments.Add(attachment);

            // Configure the SMTP client. Adjust host, port, and credentials as needed.
            SmtpClient smtp = new SmtpClient("smtp.example.com", 587)
            {
                EnableSsl = true,
                Credentials = new NetworkCredential("smtp_user", "smtp_password")
            };

            // Send the email.
            smtp.Send(message);
        }
    }
}
