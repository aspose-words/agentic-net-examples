using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontSubstitutionPdfRenderer
{
    static void Main()
    {
        // Input and output directories (adjust paths as needed).
        string inputDir = @"C:\Docs\Input\";
        string outputDir = @"C:\Docs\Output\";

        // Load the source document.
        Document doc = new Document(inputDir + "Rendering.docx");

        // Collect font substitution warnings during loading/saving.
        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.WarningCallback = warnings;

        // Configure font settings with substitution rules.
        FontSettings fontSettings = new FontSettings();
        // Use Arial as the default substitute for missing fonts.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        // If Arial is unavailable, substitute with Arvo, then Slab.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");
        // Optional: specify a folder that contains custom fonts.
        string fontsFolder = @"C:\Docs\Fonts\";
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // Save the document as PDF with full font embedding (optional).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };
        doc.Save(outputDir + "FontSubstitutionOverview.pdf", pdfOptions);

        // Output the collected font substitution warnings.
        foreach (WarningInfo info in warnings)
        {
            if (info.WarningType == WarningType.FontSubstitution)
            {
                FontSubstitutionWarningInfo fsInfo = (FontSubstitutionWarningInfo)info;
                Console.WriteLine("Font substitution warning:");
                Console.WriteLine($"  Requested Font: {fsInfo.RequestedFamilyName} (Bold={fsInfo.RequestedBold}, Italic={fsInfo.RequestedItalic})");
                Console.WriteLine($"  Resolved Font: {fsInfo.ResolvedFont.FullFontName}");
                Console.WriteLine($"  Reason: {fsInfo.Reason}");
                Console.WriteLine($"  Description: {fsInfo.Description}");
            }
        }
    }
}
