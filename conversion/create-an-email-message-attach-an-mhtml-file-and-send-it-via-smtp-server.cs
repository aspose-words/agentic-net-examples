using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to be saved as MHTML.");

        // Save the document as MHTML.
        string mhtmlPath = "sample.mhtml";
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Prepare the email message.
        using MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));
        message.Subject = "Sample MHTML Attachment";
        message.Body = "Please find the attached MHTML file.";
        message.Attachments.Add(new Attachment(mhtmlPath));

        // Configure SmtpClient to write the email to a local pickup directory instead of sending.
        string pickupDir = Path.Combine(Directory.GetCurrentDirectory(), "emails");
        Directory.CreateDirectory(pickupDir);

        using SmtpClient client = new SmtpClient("localhost")
        {
            DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
            PickupDirectoryLocation = pickupDir
        };

        // Send the email (writes .eml file to the pickup directory).
        client.Send(message);

        // Validate that an email file was created.
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email was not written to the pickup directory.");

        // Cleanup: optional removal of generated files.
        // File.Delete(mhtmlPath);
        // foreach (string file in emlFiles) File.Delete(file);
    }
}
