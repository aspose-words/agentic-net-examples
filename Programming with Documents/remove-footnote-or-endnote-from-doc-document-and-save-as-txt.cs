using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndSaveAsTxt
{
    static void Main()
    {
        // Path to the folder that contains the input document.
        string dataDir = @"C:\Docs\";

        // Load the existing Word document.
        Document doc = new Document(dataDir + "Input.docx");

        // Collect all footnote and endnote nodes in the document.
        // Both footnotes and endnotes are represented by the Footnote class.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote/endnote from its parent.
        foreach (Footnote note in notes)
        {
            note.Remove();
        }

        // Configure text save options (optional: exclude headers/footers).
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the modified document as plain text.
        doc.Save(dataDir + "Output.txt", saveOptions);
    }
}
