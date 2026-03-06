using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveNotesAndConvertToEpub
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Collect all footnote and endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each note node starting from the end to avoid index shifting.
        for (int i = notes.Count - 1; i >= 0; i--)
        {
            notes[i].Remove();
        }

        // Save the modified document as EPUB.
        doc.Save("OutputDocument.epub", SaveFormat.Epub);
    }
}
