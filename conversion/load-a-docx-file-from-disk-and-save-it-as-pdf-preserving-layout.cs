using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample DOCX document.");
        builder.Writeln("It will be converted to PDF while preserving layout.");
        const string docxPath = "sample.docx";
        sampleDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX file from disk.
        Document loadedDoc = new Document(docxPath);

        // Save the loaded document as PDF.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Optional: indicate success (no interactive output required).
        Console.WriteLine("DOCX successfully converted to PDF.");
    }
}
