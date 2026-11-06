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
        builder.Writeln("This is a sample DOCX document that will be converted to PDF with high image compression.");

        // Save the sample document as DOCX.
        string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure PDF save options for high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for images.
            ImageCompression = PdfImageCompression.Jpeg,
            // Lower JPEG quality results in higher compression (0-100, lower = higher compression).
            JpegQuality = 50
        };

        // Save the document as PDF with the specified options.
        string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optionally, you could also check the file size to confirm compression.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"PDF saved successfully. File size: {fileSize} bytes.");
    }
}
