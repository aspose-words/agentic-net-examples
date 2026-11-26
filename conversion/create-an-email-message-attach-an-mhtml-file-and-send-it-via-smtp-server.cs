using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as MHTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a test email with an MHTML attachment.");
        string mhtmlPath = "sample.mhtml";
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Prepare the email message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add(new MailAddress("recipient@example.com"));
            message.Subject = "Test Email with MHTML Attachment";
            message.Body = "Please find the attached MHTML file.";

            // Attach the MHTML file and ensure the attachment is disposed after sending.
            using (Attachment attachment = new Attachment(mhtmlPath))
            {
                message.Attachments.Add(attachment);

                // Configure SMTP client to use a local pickup directory (simulates sending).
                string pickupDir = Path.Combine(Directory.GetCurrentDirectory(), "mailpickup");
                Directory.CreateDirectory(pickupDir);

                using (SmtpClient client = new SmtpClient("localhost", 25))
                {
                    client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    client.PickupDirectoryLocation = pickupDir;
                    client.EnableSsl = false; // Not needed for local pickup.

                    client.Send(message);
                }

                // Verify that an email file was written to the pickup directory.
                string[] files = Directory.GetFiles(pickupDir);
                if (files.Length == 0)
                    throw new InvalidOperationException("No email was written to the pickup directory.");

                // Clean up email files.
                foreach (string file in files)
                {
                    File.Delete(file);
                }

                // Remove the pickup directory.
                Directory.Delete(pickupDir);
            }
        }

        // Clean up the MHTML file after the attachment has been disposed.
        File.Delete(mhtmlPath);
    }
}
