using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Paths to the source DOCX and the destination PDF.
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Set the default language of the document to Russian (ru‑RU).
        // This influences hyphenation and other language‑specific features.
        doc.Styles.DefaultFont.LocaleId = new CultureInfo("ru-RU").LCID;

        // Enable automatic hyphenation so that the Russian language rules are applied.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
