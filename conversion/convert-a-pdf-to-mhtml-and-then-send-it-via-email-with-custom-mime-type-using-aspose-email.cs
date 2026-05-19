using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion to MHTML.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 2: Load the PDF and convert it to MHTML.
        Document pdfDoc = new Document(pdfPath);
        const string mhtmlPath = "sample.mht";
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true // Use CID URLs for resources.
        };
        pdfDoc.Save(mhtmlPath, mhtmlOptions);
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Step 3: Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);
        byte[] mhtmlBytes = File.ReadAllBytes(mhtmlPath);

        // Step 4: Create an email message with a custom MIME type.
        MailMessage message = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "PDF converted to MHTML",
            IsBodyHtml = true,
            Body = mhtmlContent
        };
        message.To.Add(new MailAddress("recipient@example.com"));

        // Attach the MHTML file with a custom MIME type.
        Attachment attachment = new Attachment(new MemoryStream(mhtmlBytes), "sample.mht", "application/x-custom-mhtml");
        message.Attachments.Add(attachment);

        // Step 5: Save the email to an .eml file (simulating sending).
        const string emlPath = "email.eml";
        SaveMailMessageAsEml(message, emlPath);
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("EML file was not created.");

        // Optional: Attempt to send via SMTP (wrapped in try/catch to avoid runtime errors if no server is available).
        try
        {
            using SmtpClient client = new SmtpClient("smtp.example.com", 25);
            client.Send(message);
        }
        catch
        {
            // Ignored – the example focuses on creation and saving of the email.
        }
    }

    // Helper method to write a MailMessage to an .eml file using a simple MIME format.
    private static void SaveMailMessageAsEml(MailMessage message, string filePath)
    {
        var sb = new StringBuilder();

        sb.AppendLine($"From: {message.From}");
        sb.AppendLine($"To: {string.Join(", ", message.To.Select(a => a.Address))}");
        sb.AppendLine($"Subject: {message.Subject}");
        sb.AppendLine("MIME-Version: 1.0");
        sb.AppendLine("Content-Type: multipart/mixed; boundary=\"BOUNDARY\"");
        sb.AppendLine();

        // Body part (HTML)
        sb.AppendLine("--BOUNDARY");
        sb.AppendLine("Content-Type: text/html; charset=utf-8");
        sb.AppendLine();
        sb.AppendLine(message.Body);
        sb.AppendLine();

        // Attachment part
        foreach (Attachment att in message.Attachments)
        {
            sb.AppendLine("--BOUNDARY");
            sb.AppendLine($"Content-Type: {att.ContentType.MediaType}; name=\"{att.Name}\"");
            sb.AppendLine("Content-Transfer-Encoding: base64");
            sb.AppendLine();

            using var ms = new MemoryStream();
            att.ContentStream.CopyTo(ms);
            string base64 = Convert.ToBase64String(ms.ToArray());
            sb.AppendLine(base64);
            sb.AppendLine();
        }

        sb.AppendLine("--BOUNDARY--");

        File.WriteAllText(filePath, sb.ToString());
    }
}
