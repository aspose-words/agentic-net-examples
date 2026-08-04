using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF by first creating a Word document and saving it as PDF.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = "sample.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Expected output DOCX was not created.");
    }
}
