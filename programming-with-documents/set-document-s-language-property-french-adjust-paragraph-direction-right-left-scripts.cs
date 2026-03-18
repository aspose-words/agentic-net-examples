using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fonts;

class SetLanguageAndDirection
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the default language of the document to French (France).
        // This affects the locale of text that does not have an explicit language set.
        doc.Styles.DefaultFont.LocaleId = new CultureInfo("fr-FR").LCID;

        // Use a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The builder inherits the document's default language, so this run will be French.
        builder.Font.LocaleId = new CultureInfo("fr-FR").LCID;
        builder.Writeln("Bonjour le monde!");

        // Create a right‑to‑left paragraph (e.g., Hebrew).
        // Set the paragraph's Bidi property to true to indicate RTL layout.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!"); // Hebrew text displayed RTL.

        // Return to left‑to‑right for subsequent paragraphs if needed.
        builder.ParagraphFormat.Bidi = false;
        builder.Writeln("Encore du texte en français.");

        // Save the document.
        doc.Save("Document_French_RTL.docx");
    }
}
