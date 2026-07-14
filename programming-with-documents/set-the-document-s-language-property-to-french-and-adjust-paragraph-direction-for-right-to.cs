using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document's default language to French (France).
        // This changes the locale of the default font, which Word uses as the document language.
        doc.Styles.DefaultFont.LocaleId = new CultureInfo("fr-FR").LCID;

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a French paragraph.
        builder.Writeln("Bonjour le monde!");

        // Add a right‑to‑left paragraph (e.g., Hebrew).
        // Enable the Bidi flag for the paragraph to make the layout right‑to‑left.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!");

        // Reset Bidi for any following left‑to‑right paragraphs if needed.
        builder.ParagraphFormat.Bidi = false;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "Document_French_RTL.docx");
        doc.Save(outputPath);
    }
}
