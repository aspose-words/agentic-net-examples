using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontSubstitutionRender
{
    static void Main()
    {
        // Paths to input document, fonts folder and output directory.
        string MyDir = @"C:\Docs\";
        string FontsDir = @"C:\Fonts\";
        string ArtifactsDir = @"C:\Output\";

        // Load the document that contains the font‑substitution overview.
        Document doc = new Document(MyDir + "Rendering.docx");

        // Collect any warnings (e.g., font‑substitution warnings) that occur during loading/saving.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Configure font settings: set a default substitute font and add a table substitution rule.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        fontSettings.SetFontsFolder(FontsDir, false);
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Create PDF save options – embed the full fonts to preserve appearance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedFullFonts = true;

        // Save the rendered document as PDF.
        doc.Save(ArtifactsDir + "FontSubstitutionOverview.pdf", pdfOptions);

        // (Optional) Output collected font‑substitution warnings.
        foreach (WarningInfo info in warningCollector)
        {
            if (info.WarningType == WarningType.FontSubstitution)
                Console.WriteLine(info.Description);
        }
    }
}
