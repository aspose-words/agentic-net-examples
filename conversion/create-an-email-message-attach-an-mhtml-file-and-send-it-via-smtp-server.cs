using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        if (!Directory.Exists(artifactsDir))
            Directory.CreateDirectory(artifactsDir);

        string pickupDir = Path.Combine(Directory.GetCurrentDirectory(), "MailPickup");
        if (!Directory.Exists(pickupDir))
            Directory.CreateDirectory(pickupDir);

        // 1. Create a simple Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document that will be saved as MHTML.");

        // 2. Save the document as MHTML
        string mhtmlPath = Path.Combine(artifactsDir, "sample.mht");
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // 3. Create an email message and attach the MHTML file
        MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));
        message.Subject = "Sample MHTML Attachment";
        message.Body = "Please find the MHTML document attached.";
        Attachment attachment = new Attachment(mhtmlPath);
        message.Attachments.Add(attachment);

        // 4. Configure SmtpClient to use a local pickup directory (no real server needed)
        SmtpClient smtpClient = new SmtpClient("localhost");
        smtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
        smtpClient.PickupDirectoryLocation = pickupDir;

        // 5. Send the email (will be written as an .eml file in the pickup directory)
        smtpClient.Send(message);

        // 6. Verify that an .eml file was created
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email was not written to the pickup directory.");

        // Clean up resources
        attachment.Dispose();
        message.Dispose();
        smtpClient.Dispose();

        // Optional: indicate success (no console output required by specifications)
    }
}
