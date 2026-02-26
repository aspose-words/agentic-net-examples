using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes; // Added correct namespace for Footnote

class RemoveNotesAndSaveAsTxt
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Remove all footnotes and endnotes.
        // Get all footnote nodes (both footnotes and endnotes are represented by the Footnote class).
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote note in notes.Cast<Footnote>())
        {
            note.Remove();
        }

        // Configure text save options (optional: exclude headers/footers).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the modified document as plain text.
        doc.Save("Output.txt", txtOptions);
    }
}
