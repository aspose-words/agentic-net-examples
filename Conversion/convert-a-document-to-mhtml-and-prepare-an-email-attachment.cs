using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;
using Aspose.Words.Saving;

class MhtmlEmailAttachmentExample
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("InputDocument.docx");

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for embedded resources to improve compatibility with email clients.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: make the output more readable.
            PrettyFormat = true
        };

        // Save the document to a memory stream in MHTML format.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, saveOptions);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create an email message.
            MailMessage message = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Document as MHTML attachment",
                Body = "Please find the attached document."
            };
            message.To.Add("recipient@example.com");

            // Attach the MHTML stream to the email.
            // Use the appropriate MIME type for MHTML (message/rfc822) or application/octet-stream.
            Attachment attachment = new Attachment(mhtmlStream, "Document.mht", MediaTypeNames.Application.Octet);
            message.Attachments.Add(attachment);

            // Send the email using an SMTP client (configure as needed).
            using (SmtpClient smtp = new SmtpClient("smtp.example.com"))
            {
                // smtp.Credentials = new System.Net.NetworkCredential("user", "password");
                // smtp.EnableSsl = true;
                smtp.Send(message);
            }

            // Dispose of the attachment (which also disposes the stream).
            attachment.Dispose();
        }
    }
}
