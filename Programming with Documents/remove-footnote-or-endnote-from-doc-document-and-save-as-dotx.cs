using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Collect all footnote and endnote nodes (both are of type Footnote).
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote node. Iterate backwards to avoid collection changes.
        for (int i = footnotes.Count - 1; i >= 0; i--)
        {
            footnotes[i].Remove();
        }

        // Save the modified document as a DOTX template.
        doc.Save("Output.dotx", SaveFormat.Dotx);
    }
}
