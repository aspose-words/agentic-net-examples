using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will be saved as MHTML and attached to an email.");

        // Step 2: Save the document as MHTML.
        string mhtmlPath = Path.Combine(Path.GetTempPath(), "sample.mht");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save(mhtmlPath, saveOptions);

        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Step 3: Prepare the email message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress("sender@example.com");
            message.To.Add(new MailAddress("recipient@example.com"));
            message.Subject = "Test Email with MHTML Attachment";
            message.Body = "Please find the attached MHTML document.";
            message.IsBodyHtml = false;

            // Attach the MHTML file.
            using (Attachment attachment = new Attachment(mhtmlPath))
            {
                message.Attachments.Add(attachment);

                // Step 4: Configure the SMTP client to use a pickup directory (no real server needed).
                string pickupDirectory = Path.Combine(Path.GetTempPath(), "mailpickup");
                Directory.CreateDirectory(pickupDirectory);

                using (SmtpClient smtpClient = new SmtpClient())
                {
                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    smtpClient.PickupDirectoryLocation = pickupDirectory;

                    // Send the email (writes an .eml file to the pickup directory).
                    smtpClient.Send(message);
                }

                // Verify that an .eml file was created.
                string[] emlFiles = Directory.GetFiles(pickupDirectory, "*.eml");
                if (emlFiles.Length == 0)
                    throw new InvalidOperationException("Email was not written to the pickup directory.");

                // Cleanup temporary files (optional).
                foreach (string file in emlFiles)
                    File.Delete(file);
                Directory.Delete(pickupDirectory, true);
            } // Attachment disposed here, file handle released.

            // Delete the MHTML file after the attachment has been disposed.
            File.Delete(mhtmlPath);
        } // MailMessage disposed here.
    }
}
