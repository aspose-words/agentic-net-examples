using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document for DOCX to MHTML conversion.");
        const string docxPath = "sample.docx";
        sampleDoc.Save(docxPath, SaveFormat.Docx);

        // Verify DOCX creation.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The sample DOCX file was not created.");

        // Step 2: Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Step 3: Convert the document to MHTML and save it.
        const string mhtmlPath = "sample.mhtml";
        loadedDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify MHTML creation.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("The MHTML file was not created.");

        // Step 4: Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // Step 5: Create a simple .eml file with the MHTML as the HTML body.
        const string emlPath = "email.eml";
        string emlContent =
            "From: sender@example.com\r\n" +
            "To: recipient@example.com\r\n" +
            "Subject: Document converted to MHTML\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n" +
            "\r\n" +
            mhtmlContent;

        File.WriteAllText(emlPath, emlContent);

        // Verify email file creation.
        if (!File.Exists(emlPath))
            throw new InvalidOperationException("The email file was not created.");

        // All steps completed successfully.
        Console.WriteLine("DOCX converted to MHTML and embedded in email successfully.");
    }
}
