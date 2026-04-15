using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text with a footnote and an endnote.
        builder.Write("Sample text with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "This is a footnote.");
        builder.Write(" And an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "This is an endnote.");

        // Remove all footnotes while keeping endnotes.
        // Get all footnote nodes (both footnotes and endnotes) in the document.
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        // Iterate backwards because removing nodes changes the collection.
        for (int i = footnoteNodes.Count - 1; i >= 0; i--)
        {
            Footnote footnote = (Footnote)footnoteNodes[i];
            if (footnote.FootnoteType == FootnoteType.Footnote)
            {
                footnote.Remove();
            }
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
