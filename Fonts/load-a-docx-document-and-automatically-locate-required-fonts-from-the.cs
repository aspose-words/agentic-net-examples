using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string docPath = @"C:\Docs\Sample.docx";

        // Create a FontSettings instance that will use the system font sources.
        FontSettings fontSettings = new FontSettings();
        // Reset the font sources to the default, which includes all TrueType fonts installed on the OS.
        fontSettings.ResetFontSources();

        // Optional: attach a warning callback to capture any font‑substitution warnings.
        var warningCollector = new WarningInfoCollection();

        // Create LoadOptions, assign the FontSettings and the warning callback to it.
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings,
            WarningCallback = warningCollector // <-- correct property
        };

        // Load the document using the configured LoadOptions.
        Document doc = new Document(docPath, loadOptions);

        // Example: output any font substitution warnings that occurred during loading.
        foreach (WarningInfo warning in warningCollector)
        {
            if (warning.WarningType == WarningType.FontSubstitution)
                Console.WriteLine($"Font substitution warning: {warning.Description}");
        }

        // (Optional) Save the document to verify that it renders correctly.
        string outputPath = @"C:\Docs\Sample_Processed.pdf";
        doc.Save(outputPath);
    }
}
