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

        // Add a paragraph with a footnote.
        builder.Writeln("This paragraph contains a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Add a paragraph with an endnote.
        builder.Writeln("This paragraph contains an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote text.");

        // Remove all footnotes while keeping endnotes.
        // Get all footnote/endnote nodes in the document.
        NodeCollection allFootnotes = doc.GetChildNodes(NodeType.Footnote, true);
        // Iterate backwards to safely remove nodes.
        for (int i = allFootnotes.Count - 1; i >= 0; i--)
        {
            Footnote fn = (Footnote)allFootnotes[i];
            if (fn.FootnoteType == FootnoteType.Footnote)
                fn.Remove();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
