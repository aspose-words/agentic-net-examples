using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

namespace RemoveFootnotesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample content with footnotes and endnotes.
            builder.Writeln("This is a paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1 text.");

            builder.Writeln("This is a paragraph with an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1 text.");

            builder.Writeln("Another paragraph with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2 text.");

            // Remove all footnotes while keeping endnotes.
            // Get all footnote nodes (both footnotes and endnotes) in the document.
            NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);

            // Iterate backwards to safely remove nodes from the collection.
            for (int i = footnoteNodes.Count - 1; i >= 0; i--)
            {
                Footnote footnote = (Footnote)footnoteNodes[i];
                if (footnote.FootnoteType == FootnoteType.Footnote)
                {
                    footnote.Remove();
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);
        }
    }
}
