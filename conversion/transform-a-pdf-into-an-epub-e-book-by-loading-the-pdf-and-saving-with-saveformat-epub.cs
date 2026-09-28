using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document created for conversion to EPUB.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to EPUB.
        Document pdfDocument = new Document(pdfPath);
        string epubPath = "output.epub";
        pdfDocument.Save(epubPath, SaveFormat.Epub);

        // Validate that the EPUB file was created.
        if (!File.Exists(epubPath))
        {
            throw new InvalidOperationException("Expected output EPUB was not created.");
        }

        // Optional: clean up the intermediate PDF file.
        // File.Delete(pdfPath);
    }
}
