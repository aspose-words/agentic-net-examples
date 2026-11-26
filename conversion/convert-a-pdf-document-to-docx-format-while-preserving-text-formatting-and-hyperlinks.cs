using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with formatted text and a hyperlink.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // Heading with bold formatting.
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Sample Heading");
        // Normal paragraph with a hyperlink.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.InsertHyperlink("https://www.example.com", "Example Site", false);
        builder.Writeln();

        // Save the document as PDF to use as input for conversion.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF input file was not created.");

        // Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX output file was not created.");
    }
}
