using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary PDF and final DOCX.
        const string pdfPath = "sample.pdf";
        const string docxPath = "converted.docx";

        // Create a simple document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF created by Aspose.Words.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Ensure the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify the DOCX output.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Expected DOCX file was not created.");
    }
}
