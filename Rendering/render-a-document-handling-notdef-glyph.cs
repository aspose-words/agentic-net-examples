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

        // Configure font settings to substitute any missing font with a known system font.
        FontSettings fontSettings = new FontSettings();

        // Access the default font substitution rule.
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.DefaultFontName = "Arial"; // Replace with any installed font.
        defaultSubstitution.Enabled = true; // Ensure the rule is active.

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Build document content using a font that does not exist to trigger substitution.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFontThatDoesNotExist";
        builder.Writeln("This line uses a missing font and will be rendered with the default substitute.");

        // Capture any font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Save the document to PDF (or any other fixed‑page format).
        doc.Save("Output.pdf");

        // Output warnings, if any.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
