using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample text with footnotes.
        builder.Write("First paragraph. ");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");
        builder.Write(" Second paragraph. ");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");

        // Change footnote numbering style to lower‑case Roman numerals.
        doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "FootnotesLowerRoman.docx");
        doc.Save(outputPath);
    }
}
