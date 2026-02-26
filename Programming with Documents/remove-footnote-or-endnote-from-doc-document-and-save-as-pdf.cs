using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path for the resulting PDF file.
        string outputPath = "output.pdf";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Remove all footnotes and endnotes.
        // In Aspose.Words both footnotes and endnotes are represented by the Footnote class.
        // They are retrieved using NodeType.Footnote and distinguished by the FootnoteType property.
        NodeCollection notes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote note in notes)
        {
            // Remove regardless of whether it is a footnote or an endnote.
            note.Remove();
        }

        // Save the modified document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
