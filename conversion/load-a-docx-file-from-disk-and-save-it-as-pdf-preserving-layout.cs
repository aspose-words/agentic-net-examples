using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX content for conversion.");
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file from disk.
        Document doc = new Document(inputPath);

        // Save the document as PDF, preserving layout.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optionally, clean up the sample files (comment out if you want to keep them).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
