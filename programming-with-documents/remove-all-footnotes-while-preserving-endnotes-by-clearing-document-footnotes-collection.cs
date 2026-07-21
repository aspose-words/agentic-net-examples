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

        // Insert paragraphs with footnotes.
        builder.Writeln("First paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "This is footnote 1.");

        builder.Writeln("Second paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "This is footnote 2.");

        // Insert paragraphs with endnotes.
        builder.Writeln("First paragraph with an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "This is endnote 1.");

        builder.Writeln("Second paragraph with an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "This is endnote 2.");

        // Remove all footnotes while keeping endnotes intact.
        // Get all footnote/endnote nodes, iterate backwards and delete only footnotes.
        var footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = footnoteNodes.Count - 1; i >= 0; i--)
        {
            Footnote fn = (Footnote)footnoteNodes[i];
            if (fn.FootnoteType == FootnoteType.Footnote)
                fn.Remove();
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
