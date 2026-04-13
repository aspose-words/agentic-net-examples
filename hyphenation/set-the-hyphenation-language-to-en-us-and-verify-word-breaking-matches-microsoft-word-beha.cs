using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationDemo
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and thus hyphenation.
        // 200 points ≈ 2.78 inches.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.PageHeight = 842; // Default A4 height.

        // Set the language of the text to English (United States).
        int enUsLcid = new CultureInfo("en-US").LCID;
        builder.Font.LocaleId = enUsLcid;
        builder.Font.Size = 24;

        // Write a paragraph that contains a very long word which will need hyphenation.
        builder.Writeln(
            "This paragraph contains a veryveryverylongwordthatneedshyphenation to demonstrate automatic hyphenation.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // HyphenationZone must be a positive value (distance from the right margin in 1/20 point).
        // Use the default value of 360 (0.25 inch) or any other positive value.
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to PDF – hyphenation is applied during layout.
        string pdfPath = Path.Combine(outputDir, "HyphenatedDocument.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Simple verification: load the saved PDF as a Document and check that a hyphen appears in the extracted text.
        // Note: Hyphenation inserts a hyphen character in the visual layout, which is reflected in the extracted text.
        Document loadedPdf = new Document(pdfPath);
        string extractedText = loadedPdf.GetText();

        if (!extractedText.Contains("-"))
            throw new InvalidOperationException("Hyphenation was not applied; expected a hyphen in the output text.");

        // Success – no console output required.
    }
}
