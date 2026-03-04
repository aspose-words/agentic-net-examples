using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new blank DOCX document.
        Document doc = new Document();

        // Initialize FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Set the default font that will be used when a requested font is missing.
        DefaultFontSubstitutionRule defaultFontRule = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultFontRule.DefaultFontName = "Courier New";

        // Save the document to disk.
        doc.Save("DefaultFont.docx");
    }
}
