using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample document that will be converted to PDF with high image compression.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options for high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for all images.
            ImageCompression = PdfImageCompression.Jpeg,
            // Set JPEG quality to a low value to increase compression (0‑100 range).
            JpegQuality = 10
        };

        // Save the document as PDF using the configured options.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF file was not created as expected.");
    }
}
