using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontFallbackToPdf
{
    static void Main()
    {
        // Paths to the source DOCX and the destination PDF.
        string inputPath = "input.docx";
        string outputPath = "output.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure font substitution: use Arial as the fallback for missing fonts.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Preserve original font metrics after substitution.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Save the document as PDF using PdfSaveOptions (no special options needed here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(outputPath, pdfOptions);
    }
}
