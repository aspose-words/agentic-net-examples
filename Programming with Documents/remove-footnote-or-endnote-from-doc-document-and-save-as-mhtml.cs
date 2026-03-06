using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Remove every footnote and endnote node from the document.
        // Footnotes and endnotes are both represented by nodes of type Footnote.
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Node footnote in footnoteNodes)
        {
            footnote.Remove();
        }

        // Save the modified document as MHTML.
        // The SaveFormat enum value Mhtml ensures the correct output format.
        doc.Save("Output.mhtml", SaveFormat.Mhtml);
    }
}
