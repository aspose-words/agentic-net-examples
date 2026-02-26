using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveNotesAndConvert
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Collect all footnote and endnote nodes in the document.
        var notes = doc.GetChildNodes(NodeType.Footnote, true)
                       .Cast<Footnote>()
                       .ToList();

        // Remove each footnote and endnote from the document.
        foreach (Footnote note in notes)
        {
            // Both footnotes and endnotes are represented by the Footnote class.
            // The FootnoteType property distinguishes them.
            note.Remove();
        }

        // Save the modified document as EPUB.
        doc.Save("OutputDocument.epub", SaveFormat.Epub);
    }
}
