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
        // Input document path
        string inputPath = "input.docx";

        // Email configuration
        string smtpHost = "smtp.example.com";
        int smtpPort = 587;
        string smtpUser = "user@example.com";
        string smtpPass = "password";
        string fromAddress = "sender@example.com";
        string toAddress = "recipient@example.com";
        string subject = "Converted Document";
        string body = "Please find the converted PDF attached.";

        // Load the source document (uses Aspose.Words load rule)
        Document doc = new Document(inputPath);

        // Convert the document to PDF and keep it in memory
        using (MemoryStream pdfStream = new MemoryStream())
        {
            // SaveOptions for PDF (inherits FixedPageSaveOptions)
            FixedPageSaveOptions pdfOptions = new PdfSaveOptions();
            doc.Save(pdfStream, pdfOptions); // uses Aspose.Words save rule
            pdfStream.Position = 0; // reset stream for reading

            // Create the email message
            MailMessage message = new MailMessage();
            message.From = new MailAddress(fromAddress);
            message.To.Add(toAddress);
            message.Subject = subject;
            message.Body = body;

            // Attach the converted PDF
            Attachment attachment = new Attachment(pdfStream, "ConvertedDocument.pdf", "application/pdf");
            message.Attachments.Add(attachment);

            // Send the email via SMTP
            using (SmtpClient client = new SmtpClient(smtpHost, smtpPort))
            {
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(smtpUser, smtpPass);
                client.Send(message);
            }
        }
    }
}
