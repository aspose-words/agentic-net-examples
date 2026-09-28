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
        builder.Writeln("This is a sample document that will be saved as MHTML.");

        // Define paths.
        string mhtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.mhtml");
        string emailFolder = Path.Combine(Directory.GetCurrentDirectory(), "emails");

        // Save the document as MHTML.
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Ensure the email pickup directory exists.
        Directory.CreateDirectory(emailFolder);

        // Create the email message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add(new MailAddress("recipient@example.com"));
            message.Subject = "Test Email with MHTML Attachment";
            message.Body = "Please find the attached MHTML file.";

            // Attach the MHTML file.
            using (Attachment attachment = new Attachment(mhtmlPath))
            {
                message.Attachments.Add(attachment);

                // Configure the SMTP client to use a pickup directory (no real server needed).
                using (SmtpClient client = new SmtpClient())
                {
                    client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    client.PickupDirectoryLocation = emailFolder; // Must be an absolute path.

                    // Send the email (writes .eml file to the pickup directory).
                    client.Send(message);
                }
            }
        }

        // Verify that an email file was created in the pickup directory.
        string[] emlFiles = Directory.GetFiles(emailFolder, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email was not written to the pickup directory.");

        // Optionally delete generated files (comment out if you want to inspect them).
        // File.Delete(mhtmlPath);
        // foreach (string file in emlFiles) File.Delete(file);
        // Directory.Delete(emailFolder);
    }
}
