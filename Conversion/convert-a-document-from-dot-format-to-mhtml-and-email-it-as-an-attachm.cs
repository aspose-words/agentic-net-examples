using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOT template.
        string dotPath = "Template.dot";

        // Load the DOT document using LoadOptions.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Dot };
        Document doc = new Document(dotPath, loadOptions);

        // Set up save options for MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            HtmlVersion = HtmlVersion.Xhtml
        };

        // Save the document to a memory stream in MHTML format.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, saveOptions);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create the email message.
            using (MailMessage message = new MailMessage())
            {
                message.From = new MailAddress("sender@example.com");
                message.To.Add("recipient@example.com");
                message.Subject = "Converted MHTML Document";
                message.Body = "Please find the converted document attached.";

                // Attach the MHTML stream.
                using (Attachment attachment = new Attachment(mhtmlStream, "Document.mht", MediaTypeNames.Application.Octet))
                {
                    attachment.ContentDisposition.DispositionType = DispositionTypeNames.Attachment;
                    message.Attachments.Add(attachment);

                    // Configure and send via SMTP.
                    using (SmtpClient client = new SmtpClient("smtp.example.com", 587))
                    {
                        client.Credentials = new NetworkCredential("username", "password");
                        client.EnableSsl = true;
                        client.Send(message);
                    }
                }
            }
        }
    }
}
