using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words! This is a sample document.");

        string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document (simulating an existing file scenario).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Convert the document to MHTML.
        // -----------------------------------------------------------------
        string mhtmlPath = "sample.mhtml";
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        loadedDoc.Save(mhtmlPath, mhtmlOptions);

        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // -----------------------------------------------------------------
        // 4. Build a minimal RFC‑822 email message and embed the MHTML
        //    content as the HTML body.
        // -----------------------------------------------------------------
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        string emlPath = "email.eml";

        // Simple email headers followed by a blank line and the HTML body.
        string emlContent =
            $"From: sender@example.com\r\n" +
            $"To: receiver@example.com\r\n" +
            $"Subject: Document as MHTML\r\n" +
            $"MIME-Version: 1.0\r\n" +
            $"Content-Type: text/html; charset=utf-8\r\n" +
            $"\r\n" +
            $"{mhtmlContent}";

        File.WriteAllText(emlPath, emlContent);

        if (!File.Exists(emlPath))
            throw new InvalidOperationException("Email file was not created.");
    }
}
