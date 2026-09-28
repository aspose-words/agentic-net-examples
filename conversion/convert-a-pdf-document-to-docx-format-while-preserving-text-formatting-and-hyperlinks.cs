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

        // Add a heading with bold formatting.
        builder.Font.Size = 24;
        builder.Font.Bold = true;
        builder.Writeln("Sample PDF Document");

        // Add a paragraph with italic and colored text.
        builder.Font.Size = 12;
        builder.Font.Italic = true;
        builder.Font.Color = System.Drawing.Color.Blue; // Note: Color struct is from System.Drawing, allowed for color value only.
        builder.Writeln("This paragraph demonstrates text formatting.");

        // Insert a hyperlink.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);
        builder.Writeln();

        // Save the document as PDF (input for conversion).
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF input file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert and save as DOCX.
        string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify the DOCX output was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Expected DOCX output file was not created.");

        // Optionally, clean up temporary files (commented out if you want to inspect them).
        // File.Delete(pdfPath);
    }
}
