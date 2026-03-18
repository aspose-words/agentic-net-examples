using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some text using a font that is not installed.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Set up font settings with a default fallback font.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";
        doc.FontSettings = fontSettings;

        // Save the document as PDF. The missing font will be rendered with the fallback font.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("FallbackFontExample.pdf", pdfOptions);
    }
}
