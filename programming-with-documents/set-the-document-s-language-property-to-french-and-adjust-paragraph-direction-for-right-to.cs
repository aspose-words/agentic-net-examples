using System;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document's language to French (fr-FR) via the Normal style.
        Style normalStyle = doc.Styles["Normal"];
        if (normalStyle != null)
        {
            // LocaleId expects the Windows LCID of the culture.
            normalStyle.Font.LocaleId = new CultureInfo("fr-FR").LCID;
        }

        // Add a left‑to‑right paragraph.
        builder.Writeln("This paragraph uses the French language setting.");

        // Add a right‑to‑left paragraph (e.g., Arabic text).
        builder.ParagraphFormat.Bidi = true; // Enable right‑to‑left layout for this paragraph.
        builder.Writeln("مثال على نص من اليمين إلى اليسار");
        builder.ParagraphFormat.Bidi = false; // Reset for subsequent paragraphs.

        // Save the document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Indicate that the document was created.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
