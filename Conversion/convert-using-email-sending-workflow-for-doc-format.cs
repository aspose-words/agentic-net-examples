using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class EmailDocumentConversion
{
    static void Main()
    {
        // Paths for the source document (any supported format) and the target DOC file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";
        string targetPath = @"C:\Docs\ConvertedDocument.doc";

        // Load the source document using the Document constructor (load rule).
        Document doc = new Document(sourcePath);

        // Prepare save options to explicitly save as the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document to the target path using the Save(string, SaveOptions) overload (save rule).
        doc.Save(targetPath, saveOptions);

        // Prepare the email message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Converted DOC Document";
            message.Body = "Please find the converted DOC document attached.";

            // Attach the converted DOC file.
            message.Attachments.Add(new Attachment(targetPath));

            // Configure the SMTP client (replace with real server details).
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                smtp.EnableSsl = true;

                // Send the email.
                smtp.Send(message);
            }
        }

        // Optional: clean up the temporary file if no longer needed.
        // File.Delete(targetPath);
    }
}
