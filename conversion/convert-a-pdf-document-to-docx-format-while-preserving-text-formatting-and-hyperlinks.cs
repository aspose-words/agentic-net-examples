using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate PDF and final DOCX files.
        const string pdfPath = "sample.pdf";
        const string docxPath = "sample.docx";

        // -------------------------------------------------
        // Create a sample Word document with formatting and a hyperlink.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add bold formatted text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("This is bold text.");

        // Add italic formatted text.
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("This is italic text.");

        // Insert a hyperlink that should be preserved during conversion.
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Writeln(); // Ensure the hyperlink is on its own line.

        // -------------------------------------------------
        // Save the document as PDF (the source for conversion).
        // -------------------------------------------------
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // Load the PDF and convert it to DOCX.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Validate that the DOCX file was created.
        // -------------------------------------------------
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX output file was not created.");
    }
}
