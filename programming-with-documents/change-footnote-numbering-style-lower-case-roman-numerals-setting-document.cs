using System;
using Aspose.Words;
using Aspose.Words.Notes;

namespace FootnoteNumberingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some sample text and footnotes.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("This is a sample paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");

            builder.Writeln();
            builder.Write("Another paragraph with a second footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");

            // Change the footnote numbering style to lower‑case Roman numerals.
            // This affects all automatically numbered footnotes in the document.
            doc.FootnoteOptions.NumberStyle = NumberStyle.LowercaseRoman;

            // Save the document to a file.
            doc.Save("Footnotes_LowercaseRoman.docx");
        }
    }
}
