using System;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX file.
        Document doc = new Document("input.docx");

        // In Aspose.Words both footnotes and endnotes are represented by the Footnote node type.
        // Retrieve all Footnote nodes (which includes both footnotes and endnotes) and remove them.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            footnote.Remove();
        }

        // Save the cleaned document as Markdown.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
