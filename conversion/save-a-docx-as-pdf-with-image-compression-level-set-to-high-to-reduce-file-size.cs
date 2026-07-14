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
        builder.Writeln("Sample content for PDF conversion with high image compression.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options for high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg, // Use JPEG compression for images.
            JpegQuality = 10 // Low quality = high compression.
        };

        // Save the document as PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
