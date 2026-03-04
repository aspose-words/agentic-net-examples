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
        builder.Writeln("Hello, this is a test email body.");

        // Save the document as an MHTML file.
        string mhtmlPath = Path.Combine(Path.GetTempPath(), "email.mht");
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Create the email message.
        MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add("recipient@example.com");
        message.Subject = "Test Email with MHTML attachment";
        message.Body = "Please see the attached MHTML document.";

        // Attach the MHTML file.
        message.Attachments.Add(new Attachment(mhtmlPath));

        // Send the email via SMTP.
        using (SmtpClient client = new SmtpClient("smtp.example.com", 587))
        {
            client.Credentials = new NetworkCredential("username", "password");
            client.EnableSsl = true;
            client.Send(message);
        }

        // Clean up the temporary MHTML file.
        if (File.Exists(mhtmlPath))
            File.Delete(mhtmlPath);
    }
}
