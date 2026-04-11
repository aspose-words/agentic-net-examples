using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeWordsEmailExample");
        Directory.CreateDirectory(workDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document saved as MHTML.");

        // Save the document as MHTML.
        string mhtmlPath = Path.Combine(workDir, "Sample.mht");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save(mhtmlPath, saveOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("Failed to create the MHTML file.");

        // Create an email message.
        MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));
        message.Subject = "Aspose.Words MHTML Attachment";
        message.Body = "Please find the attached MHTML document.";

        // Attach the MHTML file.
        Attachment attachment = new Attachment(mhtmlPath, "application/octet-stream");
        message.Attachments.Add(attachment);

        // Configure the SMTP client to use a local pickup directory (no real server needed).
        string pickupDir = Path.Combine(workDir, "Pickup");
        Directory.CreateDirectory(pickupDir);
        using (SmtpClient client = new SmtpClient())
        {
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            client.PickupDirectoryLocation = pickupDir;
            client.Send(message);
        }

        // Verify that the email was written to the pickup directory.
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("The email was not saved to the pickup directory.");

        // Clean up resources.
        attachment.Dispose();
        message.Dispose();

        // Optionally, delete the temporary working directory.
        // Directory.Delete(workDir, true);
    }
}
