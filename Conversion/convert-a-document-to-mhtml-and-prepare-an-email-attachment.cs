using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Path where the MHTML file will be saved.
        string mhtmlPath = @"C:\Docs\ConvertedDocument.mht";

        // Load the Word document using the Document constructor (lifecycle rule).
        Document doc = new Document(sourceDocPath);

        // Create HtmlSaveOptions for MHTML format.
        // The constructor with SaveFormat argument follows the provided rule.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources to improve compatibility with email clients.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: make the output more readable.
            PrettyFormat = true
        };

        // Save the document as MHTML using the Save method that accepts SaveOptions.
        doc.Save(mhtmlPath, saveOptions);

        // Prepare an email with the MHTML file as an attachment.
        using (MailMessage mail = new MailMessage())
        {
            mail.From = new MailAddress("sender@example.com");
            mail.To.Add("recipient@example.com");
            mail.Subject = "Converted Document";
            mail.Body = "Please find the converted MHTML document attached.";

            // Attach the MHTML file.
            Attachment attachment = new Attachment(mhtmlPath, "application/octet-stream");
            mail.Attachments.Add(attachment);

            // Configure the SMTP client (adjust host, port, and credentials as needed).
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("smtp_user", "smtp_password");

                // Send the email.
                smtp.Send(mail);
            }
        }

        // Clean up the temporary MHTML file if it is no longer needed.
        if (File.Exists(mhtmlPath))
        {
            File.Delete(mhtmlPath);
        }
    }
}
