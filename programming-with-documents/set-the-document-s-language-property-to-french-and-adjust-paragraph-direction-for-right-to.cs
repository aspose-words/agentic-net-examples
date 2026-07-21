using System;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document's default language to French (fr-FR).
        // This affects spell checking and other language‑specific features.
        doc.Styles.DefaultFont.LocaleId = new CultureInfo("fr-FR").LCID;

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Left‑to‑right paragraph (default direction).
        builder.Writeln("Ceci est un paragraphe en français (de gauche à droite).");

        // Right‑to‑left paragraph. Enable Bidi layout for this paragraph.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("هذا نص عربي من اليمين إلى اليسار.");

        // Save the document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
