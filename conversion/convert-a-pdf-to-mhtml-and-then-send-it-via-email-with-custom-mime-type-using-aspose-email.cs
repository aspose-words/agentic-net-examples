using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths
        const string pdfPath = "sample.pdf";
        const string mhtmlPath = "sample.mhtml";
        const string emlPath = "email.eml";

        // -------------------------------------------------
        // 1. Create a simple Word document and save it as PDF
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample PDF converted to MHTML.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -------------------------------------------------
        // 2. Load the PDF and convert it to MHTML
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        pdfDoc.Save(mhtmlPath, mhtmlOptions);

        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed.");

        // -------------------------------------------------
        // 3. Prepare an email with the MHTML attached using a custom MIME type
        // -------------------------------------------------
        MailMessage message = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "PDF to MHTML conversion",
            Body = "Please find the MHTML attachment."
        };
        message.To.Add("recipient@example.com");

        // Attach the MHTML file with a custom MIME type
        Attachment attachment = new Attachment(mhtmlPath, new ContentType("application/x-custom-mhtml"));
        message.Attachments.Add(attachment);

        // -------------------------------------------------
        // 4. Save the email to an .eml file using a pickup directory (no SMTP server needed)
        // -------------------------------------------------
        string pickupDir = Path.Combine(Path.GetTempPath(), "MailPickup");
        Directory.CreateDirectory(pickupDir);

        using (SmtpClient client = new SmtpClient())
        {
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            client.PickupDirectoryLocation = pickupDir;
            client.Send(message);
        }

        // The pickup directory will contain a single .eml file; move it to the desired location
        string[] emlFiles = Directory.GetFiles(pickupDir, "*.eml");
        if (emlFiles.Length == 0)
            throw new InvalidOperationException("Email .eml file was not created.");

        File.Move(emlFiles[0], emlPath, true);

        // Verify that the .eml file was created
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("Email preparation failed.");

        // Optional cleanup of temporary files
        // File.Delete(pdfPath);
        // File.Delete(mhtmlPath);
    }
}
