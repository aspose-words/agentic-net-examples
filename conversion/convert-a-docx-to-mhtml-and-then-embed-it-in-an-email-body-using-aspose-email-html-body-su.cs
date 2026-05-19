using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        const string docxPath = "sample.docx";
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Convert the document to MHTML.
        const string mhtmlPath = "sample.mht";
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        loadedDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("MHTML conversion failed; output file is missing or empty.");

        // Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // Create an email message and embed the MHTML as the HTML body.
        MailMessage email = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "MHTML Email Example",
            IsBodyHtml = true,
            Body = mhtmlContent
        };
        email.To.Add("receiver@example.com");

        // Save the email to an .eml file (simple write of the MIME content).
        const string emlPath = "email.eml";
        File.WriteAllText(emlPath, mhtmlContent);

        // Verify that the email file was created.
        if (!File.Exists(emlPath) || new FileInfo(emlPath).Length == 0)
            throw new InvalidOperationException("Email saving failed; .eml file is missing or empty.");

        // Optional cleanup of temporary files.
        // File.Delete(docxPath);
        // File.Delete(mhtmlPath);
    }
}
