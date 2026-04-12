using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToDocxConverter
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Converted.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with formatted text and a hyperlink.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        // Add a heading with larger font size.
        builder.Font.Size = 24;
        builder.Font.Bold = true;
        builder.Writeln("Sample PDF Document");

        // Add a paragraph with normal formatting.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.Writeln("This PDF contains formatted text and a hyperlink.");

        // Insert a hyperlink field.
        builder.InsertHyperlink("https://www.example.com", "Visit Example.com", false);

        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the sample PDF file.", pdfPath);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to DOCX while preserving formatting and hyperlinks.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Conversion to DOCX failed; output file not found.", docxPath);

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine("PDF successfully converted to DOCX:");
        Console.WriteLine($"PDF path: {pdfPath}");
        Console.WriteLine($"DOCX path: {docxPath}");
    }
}
