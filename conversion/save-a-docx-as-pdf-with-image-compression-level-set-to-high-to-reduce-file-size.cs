using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOCX document created for conversion to PDF.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options to apply high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for all images.
            ImageCompression = PdfImageCompression.Jpeg,
            // Set JPEG quality to a low value (e.g., 10) for strong compression.
            JpegQuality = 10
        };

        // Save the document as PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("The expected PDF output file was not created.");
        }

        // Optionally, report success.
        Console.WriteLine("DOCX successfully converted to PDF with high image compression.");
    }
}
