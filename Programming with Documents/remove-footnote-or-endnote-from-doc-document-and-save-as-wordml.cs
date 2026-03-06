using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveFootnotesAndEndnotes
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove all footnotes and endnotes.
        // Iterate backwards to avoid modifying the collection while iterating.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = footnotes.Count - 1; i >= 0; i--)
        {
            // Each node in this collection is a Footnote (which may represent a footnote or an endnote).
            footnotes[i].Remove();
        }

        // Save the modified document as WORDML (XML) format.
        doc.Save("Output.xml", SaveFormat.WordML);
    }
}
