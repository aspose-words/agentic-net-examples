using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontSubstitutionExample
{
    // Custom warning collector to capture font substitution warnings.
    class WarningCollector : IWarningCallback
    {
        public readonly List<WarningInfo> Warnings = new List<WarningInfo>();

        public void Warning(WarningInfo info)
        {
            Warnings.Add(info);
        }
    }

    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Set up a warning collector to capture any font substitution warnings.
        WarningCollector collector = new WarningCollector();
        doc.WarningCallback = collector;

        // Create FontSettings and configure fallback and substitution.
        FontSettings fontSettings = new FontSettings();

        // Use the built‑in Microsoft Office fallback scheme (covers many Unicode ranges).
        fontSettings.FallbackSettings.LoadMsOfficeFallbackSettings();

        // Optionally, you could build a custom fallback scheme:
        // fontSettings.FallbackSettings.BuildAutomatic();

        // Set a default font to be used when no other substitute is found.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Enable the default font substitution rule (enabled by default, but set explicitly for clarity).
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;

        // Keep original font metrics so the layout remains as close as possible to the source.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document; the missing fonts will be substituted according to the settings above.
        doc.Save("Output.pdf");

        // Output any font substitution warnings that were captured.
        foreach (WarningInfo warning in collector.Warnings)
        {
            if (warning.WarningType == WarningType.FontSubstitution)
            {
                Console.WriteLine($"Warning: {warning.Description}");
            }
        }
    }
}
