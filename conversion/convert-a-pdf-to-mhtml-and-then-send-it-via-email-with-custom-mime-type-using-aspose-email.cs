using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample Word document.
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);
        builder.Writeln("This is a sample document for PDF to MHTML conversion.");

        // Step 2: Save the document as PDF.
        const string pdfPath = "sample.pdf";
        wordDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Step 3: Load the PDF and convert it to MHTML.
        Document pdfDoc = new Document(pdfPath);
        const string mhtmlPath = "sample.mht";
        pdfDoc.Save(mhtmlPath, SaveFormat.Mhtml);
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Step 4: Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // Step 5: Create a simple .eml file with custom MIME type.
        const string emlPath = "email.eml";
        using (StreamWriter writer = new StreamWriter(emlPath, false, System.Text.Encoding.UTF8))
        {
            writer.WriteLine("From: sender@example.com");
            writer.WriteLine("To: recipient@example.com");
            writer.WriteLine("Subject: Converted PDF to MHTML");
            writer.WriteLine("MIME-Version: 1.0");
            writer.WriteLine("Content-Type: application/custom-mhtml; charset=utf-8");
            writer.WriteLine(); // Blank line separates headers from body.
            writer.Write(mhtmlContent);
        }

        if (!File.Exists(emlPath))
            throw new InvalidOperationException("EML file was not created.");

        // Optional cleanup (commented out).
        // File.Delete(pdfPath);
        // File.Delete(mhtmlPath);
    }
}
