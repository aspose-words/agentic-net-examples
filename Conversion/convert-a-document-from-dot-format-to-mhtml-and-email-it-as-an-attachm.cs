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
        // Path to the source DOT (Word template) file.
        string dotPath = @"C:\Input\Template.dot";

        // Load the DOT document using the Document(string) constructor (lifecycle rule).
        Document doc = new Document(dotPath);

        // Prepare a memory stream to hold the MHTML output.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Save the document to MHTML format.
            // Use the Save(string, SaveFormat) overload (lifecycle rule) by providing a temporary file name,
            // then read the file into the stream, or directly save to the stream with SaveOptions.
            // Here we use the stream overload with SaveFormat to avoid an intermediate file.
            doc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create the email message.
            MailMessage message = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Converted MHTML Document",
                Body = "Please find the converted MHTML document attached."
            };
            message.To.Add("recipient@example.com");

            // Attach the MHTML content as a file attachment.
            // The attachment name can be any desired filename with .mhtml extension.
            Attachment attachment = new Attachment(mhtmlStream, "ConvertedDocument.mhtml", "multipart/related");
            message.Attachments.Add(attachment);

            // Configure the SMTP client (adjust host, port, and credentials as needed).
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
