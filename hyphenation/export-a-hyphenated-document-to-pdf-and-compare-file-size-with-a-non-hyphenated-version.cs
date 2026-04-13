using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the two PDF files.
        string hyphenatedPdfPath = Path.Combine(outputDir, "Hyphenated.pdf");
        string nonHyphenatedPdfPath = Path.Combine(outputDir, "NonHyphenated.pdf");

        // Create and save a hyphenated document.
        Document hyphenatedDoc = CreateSampleDocument(autoHyphenate: true);
        hyphenatedDoc.Save(hyphenatedPdfPath, SaveFormat.Pdf);

        // Create and save a non‑hyphenated document.
        Document nonHyphenatedDoc = CreateSampleDocument(autoHyphenate: false);
        nonHyphenatedDoc.Save(nonHyphenatedPdfPath, SaveFormat.Pdf);

        // Compare file sizes.
        long hyphenatedSize = new FileInfo(hyphenatedPdfPath).Length;
        long nonHyphenatedSize = new FileInfo(nonHyphenatedPdfPath).Length;
        long difference = hyphenatedSize - nonHyphenatedSize;

        Console.WriteLine($"Hyphenated PDF size: {hyphenatedSize} bytes");
        Console.WriteLine($"Non‑hyphenated PDF size: {nonHyphenatedSize} bytes");
        Console.WriteLine($"Size difference: {difference} bytes");
    }

    // Creates a sample document with a long paragraph that can be hyphenated.
    private static Document CreateSampleDocument(bool autoHyphenate)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping.
        builder.PageSetup.PageWidth = 300; // points (~4.17 inches)
        builder.PageSetup.PageHeight = 842; // A4 height in points.
        builder.PageSetup.LeftMargin = 20;
        builder.PageSetup.RightMargin = 20;

        // Sample text containing many long words.
        string sampleText = "Antidisestablishmentarianism is often cited as one of the longest words in the English language. " +
                            "Nevertheless, the quick brown fox jumps over the lazy dog while demonstrating extraordinary capabilities " +
                            "of typographic hyphenation within constrained layout environments.";

        builder.Font.Size = 12;
        builder.Writeln(sampleText);

        // Enable or disable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = autoHyphenate;
        // Optional: fine‑tune hyphenation behavior.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;

        return doc;
    }
}
