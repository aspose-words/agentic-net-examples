using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path to the output PDF file.
        string outputPath = @"C:\Docs\ResultDocument.pdf";

        // Load the document.
        Document doc = new Document(inputPath);

        // Set up a warning collector to capture font substitution warnings.
        WarningInfoCollection warningCollector = new WarningInfoCollection();
        doc.WarningCallback = warningCollector;

        // Configure font settings with a fallback (default) font.
        FontSettings fontSettings = new FontSettings();

        // Use the default font substitution rule to replace missing fonts with Arial.
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.Enabled = true;               // Ensure the rule is enabled.
        defaultSubstitution.DefaultFontName = "Arial";    // Fallback font name.

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Optional: keep original font metrics after substitution.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Prepare PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Output any font substitution warnings that occurred during loading or saving.
        foreach (WarningInfo warning in warningCollector)
        {
            if (warning.WarningType == WarningType.FontSubstitution)
                Console.WriteLine($"Font substitution warning: {warning.Description}");
        }
    }
}
