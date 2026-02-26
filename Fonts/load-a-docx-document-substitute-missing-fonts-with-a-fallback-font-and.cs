using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure font substitution: use Arial as the fallback for missing fonts.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Preserve original font metrics after substitution.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Create PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF, applying the font fallback settings.
        doc.Save(outputPath, pdfOptions);
    }
}
