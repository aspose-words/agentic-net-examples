using System;
using System.IO;
using Aspose.Words;

public class PdfToDocxConverter
{
    public static void Main()
    {
        // Define file names for the intermediate PDF and the final DOCX.
        const string pdfFileName = "sample.pdf";
        const string docxFileName = "converted.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample Word document with formatted text and a hyperlink.
        // -----------------------------------------------------------------
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);

        // Add a heading with larger font size.
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Sample PDF Document");

        // Add normal paragraph text.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.Writeln("This PDF contains a hyperlink that should be preserved after conversion.");

        // Insert a hyperlink.
        builder.InsertHyperlink("Aspose Home", "https://www.aspose.com", false);
        builder.Writeln(); // Move to the next line.

        // -----------------------------------------------------------------
        // Step 2: Save the document as PDF.
        // -----------------------------------------------------------------
        sourceDocument.Save(pdfFileName, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException($"The PDF file '{pdfFileName}' was not created.");

        // -----------------------------------------------------------------
        // Step 3: Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfFileName);
        pdfDocument.Save(docxFileName, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxFileName))
            throw new InvalidOperationException($"The DOCX file '{docxFileName}' was not created.");

        // The conversion is complete. No further action is required.
    }
}
