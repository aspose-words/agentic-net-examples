using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF file using Aspose.Words
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Step 2: Load the PDF document (no special LoadOptions required)
        Document pdfDoc = new Document("input.pdf");

        // Step 3: Convert PDF to MHTML
        pdfDoc.Save("output.mhtml", SaveFormat.Mhtml);
        if (!File.Exists("output.mhtml"))
            throw new InvalidOperationException("The MHTML file was not created.");

        // Step 4: Create an email with a custom MIME type attachment
        MailMessage message = new MailMessage(
            "sender@example.com",
            "receiver@example.com",
            "Converted PDF to MHTML",
            "Please find the converted MHTML file attached.");

        // Create a custom content type for the attachment
        ContentType customContentType = new ContentType("application/x-custom-mhtml")
        {
            Name = "converted.mhtml"
        };
        Attachment attachment = new Attachment("output.mhtml", customContentType);
        message.Attachments.Add(attachment);

        // Step 5: Simulate sending the email by saving it to a .eml file using a pickup directory
        string pickupDir = Path.Combine(Directory.GetCurrentDirectory(), "EmailPickup");
        Directory.CreateDirectory(pickupDir);

        using (SmtpClient client = new SmtpClient())
        {
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            client.PickupDirectoryLocation = pickupDir;
            client.Send(message);
        }

        // Step 6: Validate that an .eml file was created in the pickup directory
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("The email file was not created.");

        // Optional: Output the path of the generated .eml file for verification
        Console.WriteLine($"Email saved to: {emlFiles[0]}");
    }
}
