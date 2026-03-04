using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourceDocxPath = "input.docx";

        // Load the DOCX document using Aspose.Words.
        Document doc = new Document(sourceDocxPath);

        // Convert the document to MHTML and keep it in a memory stream.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Save the document as MHTML (uses the Save(string, SaveFormat) rule internally).
            doc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create the email message.
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Converted MHTML Document",
                Body = "Please find the converted document attached."
            };
            mail.To.Add("recipient@example.com");

            // Attach the MHTML content.
            Attachment attachment = new Attachment(mhtmlStream, "document.mhtml", "message/rfc822");
            mail.Attachments.Add(attachment);

            // Configure and send the email via SMTP.
            using (SmtpClient smtp = new SmtpClient("smtp.example.com", 587))
            {
                smtp.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }
    }
}
