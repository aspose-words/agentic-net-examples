using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file to act as the input.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content generated for conversion.");
        string inputPdfPath = "input.pdf";
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Load the PDF file.
        Document pdfDocument = new Document(inputPdfPath);

        // Convert and save as DOCX.
        string outputDocxPath = "output.docx";
        pdfDocument.Save(outputDocxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(outputDocxPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(inputPdfPath);
        // File.Delete(outputDocxPath);
    }
}
