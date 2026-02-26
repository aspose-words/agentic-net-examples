using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting SVG will be saved.
        string outputPath = "output.svg";

        // Load the document.
        Document doc = new Document(inputPath);

        // Remove all footnote nodes.
        foreach (Node footnote in doc.SelectNodes("//Footnote"))
            footnote.Remove();

        // Remove all endnote nodes.
        foreach (Node endnote in doc.SelectNodes("//Endnote"))
            endnote.Remove();

        // Configure SVG save options.
        SvgSaveOptions options = new SvgSaveOptions
        {
            // Render text as placed glyphs (curves) – makes the SVG selectable as an image.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Optional: remove page border and fit the SVG to the viewport.
            ShowPageBorder = false,
            FitToViewPort = true
        };

        // Save the modified document as SVG.
        doc.Save(outputPath, options);
    }
}
