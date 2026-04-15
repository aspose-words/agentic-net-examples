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

        // Add some text with footnotes.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");
        builder.Write(" More text. ");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");

        // Change footnote numbering style to lower‑case Roman numerals.
        doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FootnoteRoman.docx");
        doc.Save(outputPath);
    }
}
