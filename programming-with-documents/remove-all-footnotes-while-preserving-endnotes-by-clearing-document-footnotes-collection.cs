using System;
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

            // Insert some footnotes.
            builder.Write("This is some text with a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "First footnote.");
            builder.Write(" More text with another footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Second footnote.");

            // Insert some endnotes (these should be preserved).
            builder.Write(" This is text with an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "First endnote.");
            builder.Write(" More text with another endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "Second endnote.");

            // Remove all footnotes while keeping endnotes.
            // Get all footnote/endnote nodes in the document.
            NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
            // Iterate backwards to safely remove nodes.
            for (int i = footnoteNodes.Count - 1; i >= 0; i--)
            {
                Footnote footnote = (Footnote)footnoteNodes[i];
                if (footnote.FootnoteType == FootnoteType.Footnote)
                    footnote.Remove();
            }

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
