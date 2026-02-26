using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Collect all footnote and endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote from its parent.
        foreach (Footnote note in notes)
        {
            note.Remove();
        }

        // Save the modified document as a DOTX template.
        doc.Save("Output.dotx", SaveFormat.Dotx);
    }
}
