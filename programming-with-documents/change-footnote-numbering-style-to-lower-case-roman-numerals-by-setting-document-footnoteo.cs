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

        // Build some content with footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample paragraph with footnotes.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Third footnote.");

        // Change footnote numbering style to lower‑case Roman numerals.
        doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteNumberStyle.docx");
        doc.Save(outputPath);
    }
}
