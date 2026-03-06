using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndSaveSvg
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        // Replace with the actual path to your input file.
        string inputPath = @"C:\Docs\Input.docx";
        Document doc = new Document(inputPath);

        // Remove all footnotes and endnotes.
        // Footnotes and endnotes are stored as nodes of type Footnote.
        NodeCollection noteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote note in noteNodes)
        {
            // Remove each footnote/endnote from the document.
            note.Remove();
        }

        // Configure SVG save options (optional settings can be adjusted).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Render text as placed glyphs so the SVG contains curves instead of selectable text.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Remove any JavaScript from links in the SVG.
            RemoveJavaScriptFromLinks = true
        };

        // Save the modified document as an SVG file.
        // Replace with the desired output path.
        string outputPath = @"C:\Docs\Output.svg";
        doc.Save(outputPath, svgOptions);
    }
}
