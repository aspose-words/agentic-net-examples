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
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a test email with an MHTML attachment.");

        // Prepare output directory and file name.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string mhtmlPath = Path.Combine(outputDir, "TestDocument.mht");

        // Save the document as MHTML using the built‑in Save method (rule).
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Create the email message.
        MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add("recipient@example.com");
        message.Subject = "Test email with MHTML attachment";
        message.Body = "Please find the attached MHTML document.";

        // Attach the saved MHTML file.
        Attachment attachment = new Attachment(mhtmlPath, "application/octet-stream");
        message.Attachments.Add(attachment);

        // Send the email via SMTP.
        using (SmtpClient client = new SmtpClient("localhost"))
        {
            client.Port = 25;
            // Uncomment and set credentials if your SMTP server requires authentication.
            // client.Credentials = new NetworkCredential("username", "password");
            client.Send(message);
        }

        // Clean up resources.
        attachment.Dispose();
        message.Dispose();
    }
}
