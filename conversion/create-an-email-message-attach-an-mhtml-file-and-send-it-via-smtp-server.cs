using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will be saved as MHTML.");

        // Define the MHTML file name.
        string mhtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.mhtml");

        // Save the document as MHTML using the appropriate SaveFormat.
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Prepare a temporary folder for the SMTP pickup directory.
        string pickupDir = Path.Combine(Directory.GetCurrentDirectory(), "MailPickup");
        if (!Directory.Exists(pickupDir))
            Directory.CreateDirectory(pickupDir);

        // Create the email message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add(new MailAddress("recipient@example.com"));
            message.Subject = "Test Email with MHTML Attachment";
            message.Body = "Please find the attached MHTML document.";

            // Attach the generated MHTML file.
            Attachment attachment = new Attachment(mhtmlPath);
            message.Attachments.Add(attachment);

            // Configure the SMTP client to use the pickup directory (simulates sending).
            using (SmtpClient smtp = new SmtpClient())
            {
                smtp.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                smtp.PickupDirectoryLocation = pickupDir;

                // Send the email (writes an .eml file to the pickup directory).
                smtp.Send(message);
            }
        }

        // Validate that an .eml file was generated in the pickup directory.
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email was not written to the pickup directory.");

        // Clean up generated files (optional).
        File.Delete(mhtmlPath);
        foreach (string file in emlFiles)
            File.Delete(file);
        Directory.Delete(pickupDir);
    }
}
