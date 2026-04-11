using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample DOCX document.
        string docxPath = Path.Combine(workDir, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document created by Aspose.Words.");
        doc.Save(docxPath, SaveFormat.Docx);

        // Verify DOCX creation.
        if (!File.Exists(docxPath) || new FileInfo(docxPath).Length == 0)
            throw new InvalidOperationException("Failed to create the DOCX file.");

        // 2. Load the DOCX and convert it to MHTML.
        Document loadedDoc = new Document(docxPath);
        string mhtmlPath = Path.Combine(workDir, "Sample.mht");
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        loadedDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify MHTML creation.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("Failed to create the MHTML file.");

        // 3. Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // 4. Build a simple .eml file that contains the MHTML as the HTML body.
        // Since Aspose.Email is not available, we manually compose the RFC‑822 message.
        string emlPath = Path.Combine(workDir, "Email.eml");
        string emlContent =
            "From: sender@example.com\r\n" +
            "To: recipient@example.com\r\n" +
            "Subject: Document as MHTML\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n" +
            "\r\n" +
            mhtmlContent;

        File.WriteAllText(emlPath, emlContent);

        // Verify that the .eml file was created.
        if (!File.Exists(emlPath) || new FileInfo(emlPath).Length == 0)
            throw new InvalidOperationException("Failed to create the email file.");

        // Example completed.
    }
}
