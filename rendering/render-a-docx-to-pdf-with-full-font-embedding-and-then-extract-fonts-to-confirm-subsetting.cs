using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string docxPath = "sample.docx";
        string pdfPath = "output.pdf";

        // Create a simple DOCX document that uses a TrueType font
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use a common TrueType font
        builder.Writeln("This is a sample paragraph to test full font embedding in PDF.");

        // Save the DOCX locally
        doc.Save(docxPath);

        // Configure PDF save options to embed full fonts (disable subsetting)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the entire font file instead of creating a subset.
            EmbedFullFonts = true
        };

        // Render the DOCX to PDF with the configured options
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Read PDF content as text for inspection (ASCII is sufficient for marker strings)
        string pdfContent = File.ReadAllText(pdfPath, Encoding.ASCII);

        // Check for embedded font markers (e.g., /FontFile, /FontFile2, /FontFile3)
        bool hasEmbeddedFontMarker = pdfContent.Contains("/FontFile") ||
                                     pdfContent.Contains("/FontFile2") ||
                                     pdfContent.Contains("/FontFile3");

        if (!hasEmbeddedFontMarker)
            throw new InvalidOperationException("No embedded font markers were found in the PDF.");

        // Verify that subsetting is disabled by ensuring no subset font name prefixes (e.g., ABCDEF+)
        // Subset fonts typically have a six-letter uppercase prefix followed by '+'
        bool hasSubsetPrefix = Regex.IsMatch(pdfContent, @"[A-Z]{6}\+");

        if (hasSubsetPrefix)
            throw new InvalidOperationException("Subset font prefixes were detected; full font embedding was not applied.");

        // If we reach this point, the PDF has full font embedding without subsetting
        Console.WriteLine("PDF rendered with full font embedding and without subsetting successfully.");
    }
}
