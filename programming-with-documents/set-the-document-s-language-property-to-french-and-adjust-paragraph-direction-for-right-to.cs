using System;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default language for newly added text to French (France).
        builder.Font.LocaleId = new CultureInfo("fr-FR").LCID;

        // Write a left‑to‑right paragraph (default direction).
        builder.Writeln("Bonjour le monde!"); // French text.

        // Add a paragraph that should be displayed right‑to‑left.
        // Set the paragraph format's Bidi flag to true.
        builder.ParagraphFormat.Bidi = true;
        builder.Writeln("שלום עולם!"); // Hebrew text as an example of RTL.

        // Reset Bidi for subsequent paragraphs if needed.
        builder.ParagraphFormat.Bidi = false;

        // Save the document to a file in the same folder as the executable.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
