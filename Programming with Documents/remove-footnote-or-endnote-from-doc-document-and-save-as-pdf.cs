using System;
using Aspose.Words;
using Aspose.Words.Notes; // Needed for Footnote class

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Collect all footnote and endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote from its parent. Iterate backwards to avoid collection modification issues.
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            Footnote note = (Footnote)notes[i];
            note.Remove();
        }

        // Save the resulting document as PDF.
        doc.Save("Output.pdf");
    }
}
