using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class SubstituteMissingFonts
{
    static void Main()
    {
        // Path to the source DOCX that contains missing fonts.
        string inputPath = "MyDir\\MissingFont.docx";

        // Path where the processed document will be saved.
        string outputPath = "ArtifactsDir\\Result.pdf";

        // Optional: folder that contains additional fonts you want to make available.
        // If you have custom fonts, set this to the folder that holds them.
        string fontsFolder = "FontsDir";

        // Load the document.
        Document doc = new Document(inputPath);

        // Collect font substitution warnings (optional, useful for diagnostics).
        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.WarningCallback = warnings;

        // Create a FontSettings instance and assign it to the document.
        FontSettings fontSettings = new FontSettings();

        // If you have a custom fonts folder, add it as a font source.
        // This step is optional; Aspose.Words already includes system fonts.
        FolderFontSource folderSource = new FolderFontSource(fontsFolder, false);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

        // Load the predefined Microsoft Office fallback scheme.
        // This scheme defines which fonts to use for specific Unicode ranges
        // when the original font does not contain the required glyphs.
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.LoadMsOfficeFallbackSettings();

        // Enable font info substitution so that Aspose.Words can try to find the closest match
        // based on font metrics before falling back to the default substitution rule.
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Set a default font to be used when no other substitution rule can resolve the font.
        // This is a safety net; the fallback scheme will usually handle most cases.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Keep the original font metrics after substitution (preserves layout as much as possible).
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Apply the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document. The output format can be any supported format (PDF, DOCX, etc.).
        doc.Save(outputPath);

        // Output any font substitution warnings that were captured.
        foreach (WarningInfo info in warnings)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
