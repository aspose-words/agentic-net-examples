using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Configure font settings.
        FontSettings fontSettings = new FontSettings();

        // Use a font that has a large Unicode coverage as the default substitute.
        // This font will be used when the original font is missing or lacks a glyph.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;

        // Build an automatic fallback scheme that distributes available fonts across Unicode ranges.
        // This helps replace missing glyphs (e.g., .notdef) with appropriate fallback fonts.
        fontSettings.FallbackSettings.BuildAutomatic();

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Set up a warning collector to capture font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Use DocumentBuilder to add text that contains characters likely missing in the primary font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont"; // intentionally non‑existent font to trigger substitution

        // Write a range of Unicode characters (Cyrillic block) that may not be present in the missing font.
        for (int code = 0x0400; code <= 0x04FF; code++)
        {
            builder.Write(Char.ConvertFromUtf32(code));
        }
        builder.Writeln();

        // Save the document to PDF (or any other format you need).
        doc.Save("RenderedDocument.pdf", SaveFormat.Pdf);

        // Output any font substitution warnings that occurred during processing.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
