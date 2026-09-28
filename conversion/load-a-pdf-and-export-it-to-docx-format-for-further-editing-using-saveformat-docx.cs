using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF (input.pdf)
        Document sample = new Document();
        DocumentBuilder builder = new DocumentBuilder(sample);
        builder.Writeln("Sample PDF content generated for conversion.");
        string pdfPath = "input.pdf";
        sample.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to DOCX
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }

        // Optional cleanup (comment out if you want to inspect files)
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
