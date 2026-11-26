using System;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF document using Aspose.Words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF content generated for conversion.");
        string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 2: Load the PDF into a memory stream.
        using MemoryStream pdfStream = new MemoryStream(File.ReadAllBytes(pdfPath));
        pdfStream.Position = 0; // Ensure the stream is at the beginning.

        // Step 3: Create an email message and attach the PDF with a custom MIME type.
        using MailMessage message = new MailMessage();
        message.From = new MailAddress("sender@example.com");
        message.To.Add(new MailAddress("recipient@example.com"));
        message.Subject = "PDF converted to MHTML";
        message.Body = "<html><body><h1>PDF to MHTML Example</h1><p>Please see the attached PDF.</p></body></html>";
        message.IsBodyHtml = true;

        // Define a custom MIME type for the attachment.
        ContentType customPdfContentType = new ContentType("application/custom-pdf")
        {
            Name = "sample.pdf"
        };

        // Create the attachment using the custom content type.
        Attachment attachment = new Attachment(pdfStream, customPdfContentType);
        message.Attachments.Add(attachment);

        // Step 4: Simulate sending the email using SmtpClient with a pickup directory.
        // This avoids real network communication.
        string pickupDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Pickup");
        Directory.CreateDirectory(pickupDirectory);

        using SmtpClient client = new SmtpClient
        {
            DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory,
            PickupDirectoryLocation = pickupDirectory
        };

        client.Send(message);

        // Verify that the email was written to the pickup directory.
        string[] sentFiles = Directory.GetFiles(pickupDirectory, "*.eml");
        if (sentFiles.Length == 0)
            throw new InvalidOperationException("Email was not written to the pickup directory.");

        // Cleanup: dispose resources (handled by using statements).
        // Verify that the PDF and generated email files exist.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file missing after processing.");

        // Optionally, you could delete the temporary files here.
    }
}
