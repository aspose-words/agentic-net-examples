using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that uses a font which is not present on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line is formatted with a missing font and will be substituted.");

        // Configure font substitution settings.
        FontSettings fontSettings = new FontSettings();

        // Map the missing font to a list of fallback fonts.
        // Aspose.Words will try "Arial" first, then "Courier New" if the previous one is unavailable.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "MissingFont", new[] { "Arial", "Courier New" });

        // Optionally set a default substitution rule for any other missing fonts.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Apply the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Alternatively, replace the missing font directly (no specific rule exists for ReplaceFont).
        // doc.ReplaceFont("MissingFont", "Arial");

        // Prepare PDF save options – embed full fonts so the resulting PDF can be edited later.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as a PDF file.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
