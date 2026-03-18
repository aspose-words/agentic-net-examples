using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize font settings.
        FontSettings fontSettings = new FontSettings();

        // Set the default substitution font to Arial Unicode MS.
        // This font contains a very large Unicode range and will be used
        // when the original font lacks glyphs for a character.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Keep original font metrics after substitution (optional but recommended).
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Build sample content using a font that likely does not have the required glyphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("Sample text with Unicode characters: Привет, 你好, 😀");

        // Save the document. The missing glyphs will be rendered using Arial Unicode MS.
        doc.Save("FallbackExample.pdf");
    }
}
