using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with formatted text and a hyperlink.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Apply some formatting.
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("This is a bold heading.");

        // Normal paragraph with a hyperlink.
        builder.Font.Bold = false;
        builder.Font.Underline = Underline.Single;
        builder.Font.Color = System.Drawing.Color.Blue; // Color is allowed via System.Drawing for simple usage.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);
        builder.Writeln(); // Move to next line.

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");
    }
}
