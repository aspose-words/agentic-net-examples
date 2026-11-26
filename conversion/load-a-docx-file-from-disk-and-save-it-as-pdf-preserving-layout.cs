using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file as input.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content for conversion.");
        string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file from disk.
        Document doc = new Document(inputPath);

        // Save the document as PDF, preserving layout.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optionally, clean up the sample input file.
        // File.Delete(inputPath);
    }
}
