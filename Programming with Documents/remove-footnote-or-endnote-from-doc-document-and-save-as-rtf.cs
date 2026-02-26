using System;
using Aspose.Words;
using Aspose.Words.Notes;

class RemoveNotes
{
    static void Main()
    {
        // Load the source DOC document.
        string dataDir = @"C:\Docs\";
        Document doc = new Document(dataDir + "input.docx");

        // Collect all footnote and endnote nodes in the document.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote from its parent.
        foreach (Footnote note in notes)
        {
            note.Remove();
        }

        // Save the modified document as RTF.
        doc.Save(dataDir + "output.rtf", SaveFormat.Rtf);
    }
}
