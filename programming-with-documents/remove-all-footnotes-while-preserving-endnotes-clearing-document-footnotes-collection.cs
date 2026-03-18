using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Get all footnote/endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove only the footnotes (preserve endnotes).
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            Footnote note = (Footnote)notes[i];
            if (note.FootnoteType == FootnoteType.Footnote)
                note.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
