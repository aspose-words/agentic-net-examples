using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("input.docx");

        // Collect all footnote and endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote from its parent.
        foreach (Footnote note in notes)
        {
            note.Remove();
        }

        // Save the modified document as Markdown.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
