using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontSubstitutionExample
{
    static void Main()
    {
        // Paths to the input DOCX and the output file.
        string dataDir = @"C:\Data\";
        string inputPath = System.IO.Path.Combine(dataDir, "MissingFont.docx");
        string outputPath = System.IO.Path.Combine(dataDir, "Result.pdf");

        // Load the document that contains missing fonts.
        Document doc = new Document(inputPath);

        // Prepare a warning collector to capture font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Create and assign FontSettings to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load the predefined Microsoft Office fallback scheme.
        // This scheme defines which fonts to use for specific Unicode ranges when the original font lacks glyphs.
        fontSettings.FallbackSettings.LoadMsOfficeFallbackSettings();

        // Set a default font to be used when no other substitution rule can resolve the missing font.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Enable font‑info substitution so Aspose.Words can find the closest match based on font metrics.
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Preserve the original font metrics after substitution (optional but often desired).
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Save the document. The chosen format can be PDF, DOCX, etc.
        doc.Save(outputPath);

        // Output any font substitution warnings that were generated during processing.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
