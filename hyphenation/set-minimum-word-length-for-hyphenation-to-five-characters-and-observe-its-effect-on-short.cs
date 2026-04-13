using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and enable hyphenation.
        // The width is set in points (1 point = 1/72 inch). 200 points ≈ 2.78 inches.
        doc.FirstSection.PageSetup.PageWidth = 200;

        // Enable automatic hyphenation for the document.
        // Aspose.Words automatically skips hyphenation for words shorter than the
        // default minimum length (5 characters). This satisfies the requirement.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write sample text containing short words (≤4 characters) and longer words (>5 characters).
        // Short words will not be hyphenated, while longer words may be split across lines.
        builder.Font.Size = 12;
        builder.Writeln(
            "Hyphenation demo: extraordinary extraordinary test abcde abcdefghij.");

        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document in DOCX format.
        string docPath = Path.Combine(outputDir, "HyphenationMinWordLength.docx");
        doc.Save(docPath);

        // Also save as PDF to visually inspect hyphenation.
        string pdfPath = Path.Combine(outputDir, "HyphenationMinWordLength.pdf");
        doc.Save(pdfPath);
    }
}
