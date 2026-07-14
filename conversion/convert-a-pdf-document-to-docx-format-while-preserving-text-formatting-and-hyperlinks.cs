using System;
using System.IO;
using Aspose.Words;

public class PdfToDocxConverter
{
    public static void Main()
    {
        // Define file names for the sample PDF and the resulting DOCX.
        const string pdfPath = "sample.pdf";
        const string docxPath = "converted.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with formatted text and a hyperlink.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        // Add a heading with bold formatting.
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Sample PDF Document");

        // Add a paragraph with a hyperlink.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);
        builder.Writeln(); // Move to the next line.

        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"The PDF file '{pdfPath}' was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to DOCX, preserving formatting and hyperlinks.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException($"The DOCX file '{docxPath}' was not created.");

        // Optional: Inform the user that conversion succeeded.
        Console.WriteLine($"Conversion completed successfully. DOCX saved to '{docxPath}'.");
    }
}
