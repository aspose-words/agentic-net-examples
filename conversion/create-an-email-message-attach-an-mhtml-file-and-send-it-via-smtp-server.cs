using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document that will be attached as MHTML.");

        // Save the document as MHTML.
        string mhtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.mht");
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Prepare the email message.
        MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));
        message.Subject = "Sample Email with MHTML Attachment";
        message.Body = "Please find the attached MHTML document.";

        // Attach the MHTML file.
        Attachment attachment = new Attachment(mhtmlPath);
        message.Attachments.Add(attachment);

        // Configure the SMTP client to use a local pickup directory (no external server required).
        string pickupDirectory = Path.Combine(Directory.GetCurrentDirectory(), "emails");
        Directory.CreateDirectory(pickupDirectory);

        using (SmtpClient smtpClient = new SmtpClient())
        {
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            smtpClient.PickupDirectoryLocation = pickupDirectory;

            // Send the email (it will be saved as an .eml file in the pickup directory).
            smtpClient.Send(message);
        }

        // Verify that an .eml file was created.
        string[] emlFiles = Directory.GetFiles(pickupDirectory, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email was not saved to the pickup directory.");

        // Clean up resources.
        attachment.Dispose();
        message.Dispose();
    }
}
