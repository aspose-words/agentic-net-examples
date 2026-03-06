using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Set up font substitution: use Arial when a font is missing.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Preserve original font metrics after substitution (optional but improves layout).
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save("Output.pdf");
    }
}
