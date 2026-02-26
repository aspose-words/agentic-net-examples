using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document doc = new Document("Input.doc");

        // Get all footnote/endnote nodes. In Aspose.Words both footnotes and endnotes are of type NodeType.Footnote.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove them in reverse order to avoid collection modification issues.
        for (int i = footnotes.Count - 1; i >= 0; i--)
        {
            footnotes[i].Remove();
        }

        // Save the modified document as DOCX.
        doc.Save("Output.docx");
    }
}
