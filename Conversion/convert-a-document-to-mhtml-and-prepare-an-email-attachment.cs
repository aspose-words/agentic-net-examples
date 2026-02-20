using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;
using Aspose.Words.Saving;

public class MhtmlEmailHelper
{
    // Converts a Word document to MHTML and returns it as a Mail attachment.
    public static Attachment ConvertToMhtmlAttachment(string sourceDocPath, string attachmentName)
    {
        // Load the source document.
        Document doc = new Document(sourceDocPath);

        // Configure MHTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Embed resources (images, fonts, CSS) using CID URLs so the MHTML is self‑contained.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: set the encoding to UTF‑8 without BOM.
            Encoding = new System.Text.UTF8Encoding(false)
        };

        // Save the document to a memory stream using the configured options.
        MemoryStream mhtmlStream = new MemoryStream();
        doc.Save(mhtmlStream, saveOptions);
        mhtmlStream.Position = 0; // Reset stream position for reading.

        // Create the email attachment from the MHTML stream.
        // The content type for MHTML is "multipart/related".
        Attachment attachment = new Attachment(mhtmlStream, attachmentName, "multipart/related");
        // Ensure the stream stays at the beginning when the attachment is read.
        attachment.ContentStream.Position = 0;
        return attachment;
    }

    // Example usage.
    public static void Main()
    {
        // Path to the source .docx file.
        string docPath = @"C:\Docs\Sample.docx";

        // Desired name for the attachment (including extension).
        string attachmentFileName = "Sample.mhtml";

        // Convert and obtain the attachment.
        Attachment mhtmlAttachment = ConvertToMhtmlAttachment(docPath, attachmentFileName);

        // Create a simple email message.
        MailMessage message = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "Document as MHTML",
            Body = "Please find the document attached."
        };
        message.To.Add("recipient@example.com");
        message.Attachments.Add(mhtmlAttachment);

        // Configure SMTP client (replace with real credentials and host).
        using (SmtpClient smtp = new SmtpClient("smtp.example.com"))
        {
            smtp.Port = 587;
            smtp.EnableSsl = true;
            // smtp.Credentials = new System.Net.NetworkCredential("user", "password");
            // smtp.Send(message);
        }

        // Dispose resources.
        mhtmlAttachment.Dispose();
        message.Dispose();
    }
}
