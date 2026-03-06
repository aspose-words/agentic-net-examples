using System;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveNotesAndConvertToHtml
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.docx");

        // Remove all footnotes and endnotes from the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote note in notes)
        {
            note.Remove();
        }

        // Save the modified document as HTML.
        doc.Save("output.html", SaveFormat.Html);
    }
}
