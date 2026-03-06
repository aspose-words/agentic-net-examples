using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Get all footnote/endnote nodes.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Iterate backwards because removing a node changes the collection.
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            Footnote note = (Footnote)notes[i];
            // Remove both footnotes and endnotes.
            if (note.FootnoteType == FootnoteType.Footnote ||
                note.FootnoteType == FootnoteType.Endnote)
            {
                note.Remove();
            }
        }

        // Save the modified document as DOC.
        doc.Save("Output.doc");
    }
}
