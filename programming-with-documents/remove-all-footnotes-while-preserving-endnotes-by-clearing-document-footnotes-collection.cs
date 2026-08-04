using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeWordsFootnoteRemoval
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add paragraphs with footnotes.
            builder.Writeln("First paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1 content.");

            builder.Writeln("Second paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2 content.");

            // Add paragraphs with endnotes.
            builder.Writeln("First paragraph with an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1 content.");

            builder.Writeln("Second paragraph with an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2 content.");

            // Remove all footnotes while preserving endnotes.
            var footnotes = doc.GetChildNodes(NodeType.Footnote, true)
                               .Cast<Footnote>()
                               .Where(fn => fn.FootnoteType == FootnoteType.Footnote)
                               .ToList();

            foreach (var footnote in footnotes)
                footnote.Remove();

            // Save the modified document.
            string outputFile = "Result.docx";
            doc.Save(outputFile);
        }
    }
}
