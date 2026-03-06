using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document doc = new Document("Input.doc");

        // Get all footnote/endnote nodes. In Aspose.Words a footnote and an endnote are both represented by the Footnote class.
        // Remove each of them from the document.
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true)
                           .OfType<Footnote>()
                           .ToList(); // materialize to avoid collection modification issues
        foreach (Footnote fn in footnotes)
        {
            fn.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save("Output.docx");
    }
}
