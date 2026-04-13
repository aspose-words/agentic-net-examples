using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and possible hyphenation.
        builder.PageSetup.PageWidth = 200; // points
        builder.PageSetup.LeftMargin = 20;
        builder.PageSetup.RightMargin = 20;

        // Set a readable font size.
        builder.Font.Size = 12;

        // First paragraph – hyphenation will be allowed.
        string longText1 = "Antidisestablishmentarianism is a long word that may be hyphenated when the line is narrow.";
        builder.Writeln(longText1);
        Paragraph paraHyphenated = doc.FirstSection.Body.FirstParagraph;

        // Second paragraph – we will suppress hyphenation for this range.
        string longText2 = "This paragraph will not be hyphenated even though it is also long and may need hyphenation.";
        builder.Writeln(longText2);
        Paragraph paraSuppressed = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        paraSuppressed.ParagraphFormat.SuppressAutoHyphens = true; // Disable hyphenation for this paragraph.

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Save the document to PDF to visualize hyphenation.
        string outPath = Path.Combine(outputDir, "HyphenationRange.pdf");
        doc.Save(outPath);

        // Validate that the output file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("Failed to create the output PDF file.");

        // Validate that the suppression flag is correctly set.
        if (!paraSuppressed.ParagraphFormat.SuppressAutoHyphens)
            throw new InvalidOperationException("SuppressAutoHyphens flag was not set as expected.");

        // Inform the user (optional, not required for verification).
        Console.WriteLine($"Document saved to: {outPath}");
    }
}
