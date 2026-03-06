using System;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveNotesExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // In Aspose.Words both footnotes and endnotes are represented by the Footnote node type.
        // Collect all footnote (including endnote) nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each note from its parent. Iterate backwards to avoid index shifting after removal.
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            notes[i].Remove();
        }

        // Save the modified document as a DOTM (Word macro‑enabled template) file.
        doc.Save("OutputDocument.dotm", SaveFormat.Dotm);
    }
}
