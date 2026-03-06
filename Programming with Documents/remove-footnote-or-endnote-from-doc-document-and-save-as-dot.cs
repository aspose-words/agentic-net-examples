using System;
using Aspose.Words;

class RemoveNotesAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // All footnotes and endnotes are represented by the Footnote node type.
        // Retrieve the collection and remove the nodes in reverse order to avoid
        // collection modification issues.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            notes[i].Remove();
        }

        // Save the modified document as a DOT template.
        doc.Save("OutputTemplate.dot", SaveFormat.Dot);
    }
}
