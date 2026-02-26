using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndSaveAsXps
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\Input.docx";

        // Load the document from file.
        Document doc = new Document(inputPath);

        // Collect all footnote/endnote nodes in the document.
        NodeCollection noteNodes = doc.GetChildNodes(NodeType.Footnote, true);

        // Remove each footnote and endnote.
        foreach (Footnote note in noteNodes)
        {
            // Both footnotes and endnotes are represented by the Footnote class.
            // The FootnoteType property tells us which kind it is.
            if (note.FootnoteType == FootnoteType.Footnote ||
                note.FootnoteType == FootnoteType.Endnote)
            {
                note.Remove();
            }
        }

        // Create XPS save options.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the modified document as XPS.
        string outputPath = @"C:\Docs\Output.xps";
        doc.Save(outputPath, xpsOptions);
    }
}
