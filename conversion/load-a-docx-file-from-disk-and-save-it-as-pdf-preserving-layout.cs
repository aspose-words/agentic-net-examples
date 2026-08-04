using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX content for conversion to PDF.");
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from disk.
        Document doc = new Document(inputPath);

        // Save the loaded document as PDF, preserving layout.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
