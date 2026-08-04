using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content and footnotes.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This is a sample paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");
        builder.Write(" Adding more text and another footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");

        // Change the footnote numbering style to lower‑case Roman numerals.
        doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

        // Save the document to the local file system.
        doc.Save("FootnoteNumberStyle.docx");
    }
}
