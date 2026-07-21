using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string docxPath = "sample.docx";
        const string mhtmlPath = "sample.mht";
        const string emlPath = "email.eml";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document for MHTML conversion.");
        // Save the DOCX so it can be loaded again (bootstrap rule).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX and convert it to MHTML.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources to make them embeddable in email bodies.
            ExportCidUrlsForMhtmlResources = true,
            // Optional: make the output more readable.
            PrettyFormat = true
        };
        loadedDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML conversion failed; file not created.");

        // -----------------------------------------------------------------
        // 3. Read the MHTML content.
        // -----------------------------------------------------------------
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // -----------------------------------------------------------------
        // 4. Create an email message and embed the MHTML as the HTML body.
        // -----------------------------------------------------------------
        MailMessage email = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "Document converted to MHTML",
            Body = mhtmlContent,
            IsBodyHtml = true
        };
        email.To.Add("recipient@example.com");

        // Save the email to an .eml file.
        // Since Aspose.Email is not available, write the raw MHTML content as the email body.
        // This creates a simple .eml file containing the necessary headers and body.
        using (StreamWriter writer = new StreamWriter(emlPath, false))
        {
            writer.WriteLine("From: {0}", email.From);
            writer.WriteLine("To: {0}", string.Join(", ", email.To));
            writer.WriteLine("Subject: {0}", email.Subject);
            writer.WriteLine("MIME-Version: 1.0");
            writer.WriteLine("Content-Type: multipart/related; boundary=\"----=_Part_0_123456.789\"");
            writer.WriteLine();
            writer.WriteLine("------=_Part_0_123456.789");
            writer.WriteLine("Content-Type: text/html; charset=\"utf-8\"");
            writer.WriteLine("Content-Transfer-Encoding: 8bit");
            writer.WriteLine();
            writer.WriteLine(mhtmlContent);
            writer.WriteLine("------=_Part_0_123456.789--");
        }

        // Verify that the email file was created.
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("Email file was not created.");
    }
}
