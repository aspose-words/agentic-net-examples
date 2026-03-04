using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime; // Added for ContentType
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertAndEmailMhtml
{
    static void Main()
    {
        // Load the source Word document.
        // Uses the Document(string) constructor rule.
        Document doc = new Document("InputDocument.docx");

        // Prepare save options for MHTML format.
        // Uses HtmlSaveOptions(SaveFormat) constructor rule.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Optional: use CID URLs for resources to improve compatibility.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: embed document properties.
            ExportDocumentProperties = true,
            // Optional: pretty format for readability.
            PrettyFormat = true
        };

        // Save the document to a memory stream in MHTML format.
        // Uses Document.Save(Stream, SaveOptions) rule.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, saveOptions);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create the email message.
            MailMessage message = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Converted MHTML Document",
                Body = "Please find the converted MHTML document attached."
            };
            message.To.Add("recipient@example.com");

            // Attach the MHTML content.
            // The content type "multipart/related" is appropriate for MHTML.
            ContentType mhtmlContentType = new ContentType("multipart/related")
            {
                Name = "Document.mht"
            };
            Attachment attachment = new Attachment(mhtmlStream, mhtmlContentType);
            message.Attachments.Add(attachment);

            // Configure the SMTP client.
            // Adjust host, port, and credentials as needed for your environment.
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");

                // Send the email.
                smtp.Send(message);
            }
        }
    }
}
