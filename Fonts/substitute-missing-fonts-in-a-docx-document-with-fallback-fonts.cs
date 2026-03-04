using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

namespace FontSubstitutionExample
{
    // Collects warnings emitted by Aspose.Words during loading/saving.
    public class WarningInfoCollector : IWarningCallback
    {
        public void Warning(WarningInfo info)
        {
            // Print only font substitution warnings.
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine($"Warning: {info.Description}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths – adjust these to your environment.
            string docsDir = @"C:\Docs\";
            string fontsDir = @"C:\MyFonts\";
            string inputPath = Path.Combine(docsDir, "MissingFont.docx");
            string outputPath = Path.Combine(docsDir, "MissingFont_Substituted.pdf");

            // Load the document.
            Document doc = new Document(inputPath);

            // Assign a warning callback to capture font substitution warnings.
            doc.WarningCallback = new WarningInfoCollector();

            // Create and configure FontSettings.
            FontSettings fontSettings = new FontSettings();

            // Use a custom folder as the only font source (optional – can be omitted to use system fonts).
            // NOTE: The constructor takes (folderPath, isRecursive). The parameter name is not "recursive".
            FolderFontSource folderSource = new FolderFontSource(fontsDir, true);
            fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

            // Load the predefined Microsoft Office fallback scheme.
            fontSettings.FallbackSettings.LoadMsOfficeFallbackSettings();

            // Enable font info substitution (finds the closest match based on font metrics).
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

            // Set a default font to be used when no other substitute is found.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

            // Keep original font metrics after substitution (preserves layout as much as possible).
            doc.LayoutOptions.KeepOriginalFontMetrics = true;

            // Apply the configured FontSettings to the document.
            doc.FontSettings = fontSettings;

            // Save the document – PDF format demonstrates the substitution result.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
