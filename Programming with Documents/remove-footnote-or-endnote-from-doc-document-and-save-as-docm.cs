using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveNotesExample
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("InputDocument.doc");

        // Remove all footnotes and endnotes. In Aspose.Words both are represented by the Footnote node type.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);
        // Iterate backwards to avoid collection modification issues.
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            notes[i].Remove();
        }

        // Save the modified document as a macro‑enabled DOCM file.
        doc.Save("OutputDocument.docm", SaveFormat.Docm);
    }
}
