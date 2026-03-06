using Aspose.Words;
using Aspose.Words.Notes;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Collect all footnote and endnote nodes.
        NodeCollection allNotes = doc.GetChildNodes(NodeType.Footnote, true);
        List<Footnote> notesToRemove = new List<Footnote>();

        foreach (Footnote note in allNotes)
        {
            // Both footnotes and endnotes are represented by the Footnote class.
            if (note.FootnoteType == FootnoteType.Footnote || note.FootnoteType == FootnoteType.Endnote)
                notesToRemove.Add(note);
        }

        // Remove the collected notes from the document.
        foreach (Footnote note in notesToRemove)
            note.Remove();

        // Save the modified document as RTF.
        doc.Save("Output.rtf", SaveFormat.Rtf);
    }
}
