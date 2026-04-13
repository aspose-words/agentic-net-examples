using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationToPdf
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "HyphenatedDocument.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and hyphenation.
        builder.PageSetup.PageWidth = 200; // points (~2.78 inches)
        builder.PageSetup.PageHeight = 842; // A4 height in points.

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: limit consecutive hyphenated lines.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // HyphenationZone must be a non‑negative value; using the default (360) avoids the exception.
        // doc.HyphenationOptions.HyphenationZone = 360; // default, can be omitted.

        // Write a long paragraph that will require hyphenation.
        builder.Font.Size = 12;
        builder.Writeln(
            "Aspose.Words provides powerful APIs for processing documents. " +
            "When a paragraph is too long to fit within the page margins, " +
            "automatic hyphenation can split words across lines without displaying explicit hyphen marks in the PDF output.");

        // Save the document as PDF. Hyphenation will be applied, but hyphen marks are not rendered in the PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the PDF file with hyphenation.");
        }

        // Optional console output.
        Console.WriteLine($"PDF created at: {pdfPath}");
    }
}
