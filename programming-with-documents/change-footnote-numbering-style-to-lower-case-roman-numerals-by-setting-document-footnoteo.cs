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

        // Use DocumentBuilder to add some text and footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample sentence with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");
        builder.Write(" Another sentence with a second footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");

        // Change the footnote numbering style to lower‑case Roman numerals.
        doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FootnoteNumberStyle.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
