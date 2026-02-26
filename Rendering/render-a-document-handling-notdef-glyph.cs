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

        // Configure font settings to substitute missing fonts with a font that has extensive Unicode coverage.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";

        // Load a predefined fallback scheme (Google Noto) to improve glyph coverage.
        fontSettings.FallbackSettings.LoadNotoFallbackSettings();

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Set up a warning collector to capture font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Build document content using a font that does not exist on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";

        // Insert a character that is likely missing in most fonts (e.g., a Telugu glyph).
        builder.Writeln("Testing missing glyph: \u0C00");

        // Save the document to PDF. The default substitution and fallback settings will handle the .notdef glyph.
        doc.Save("Output.pdf");

        // Output any font substitution warnings to the console.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
