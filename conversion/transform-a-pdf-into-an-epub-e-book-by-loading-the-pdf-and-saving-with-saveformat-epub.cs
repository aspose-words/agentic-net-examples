using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted from PDF to EPUB.");

        // Save the document as PDF (input for conversion).
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF file.
        Document pdfDocument = new Document(pdfPath);

        // Convert and save as EPUB.
        string epubPath = "output.epub";
        pdfDocument.Save(epubPath, SaveFormat.Epub);

        // Validate that the EPUB file was created.
        if (!File.Exists(epubPath))
        {
            throw new InvalidOperationException("Expected output EPUB was not created.");
        }

        // Optional cleanup (comment out if you want to keep the files).
        // File.Delete(pdfPath);
        // File.Delete(epubPath);
    }
}
