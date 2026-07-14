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
        builder.Writeln("Sample DOCX content.");
        string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(inputPath);

        // Save as PDF preserving layout.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
