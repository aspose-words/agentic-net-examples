using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF (input for conversion).
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample PDF content generated for conversion to EPUB.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Convert the PDF to EPUB.
        string epubPath = "output.epub";
        pdfDocument.Save(epubPath, SaveFormat.Epub);

        // Validate that the EPUB file was created.
        if (!File.Exists(epubPath))
        {
            throw new InvalidOperationException("Expected output EPUB was not created.");
        }

        // Optional cleanup of the temporary PDF file.
        if (File.Exists(pdfPath))
        {
            File.Delete(pdfPath);
        }
    }
}
