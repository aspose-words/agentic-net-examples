using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the default language of the document to French (fr-FR).
        // This affects the language used for spell checking and other language‑specific features.
        doc.Styles.DefaultFont.LocaleId = new CultureInfo("fr-FR").LCID;

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Left‑to‑right paragraph (default direction).
        builder.Writeln("Hello world!");

        // Right‑to‑left paragraph. Set the paragraph format's Bidi property to true.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!"); // Hebrew text displayed right‑to‑left.

        // Reset Bidi for subsequent paragraphs if needed.
        builder.ParagraphFormat.Bidi = false;

        // Save the document to a file in the same folder as the executable.
        doc.Save("Result.docx");
    }
}
