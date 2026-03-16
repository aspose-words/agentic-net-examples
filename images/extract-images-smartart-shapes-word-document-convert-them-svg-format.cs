using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class SmartArtToSvgExtractor
{
    static void Main()
    {
        // Load the source Word document.
        // Replace "Input.docx" with the path to your document containing SmartArt.
        Document doc = new Document("Input.docx");

        // Get all Shape nodes in the document (including those inside groups) that contain SmartArt.
        var smartArtShapes = doc.GetChildNodes(NodeType.Shape, true)
                               .OfType<Shape>()
                               .Where(s => s.HasSmartArt)
                               .ToList();

        // Prepare SVG save options.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Render text as placed glyphs to avoid selectable text in the SVG.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Do not add a page border around the SVG.
            ShowPageBorder = false,
            // Fit the SVG to the viewport (100% width/height).
            FitToViewPort = true
        };

        int index = 0;
        foreach (Shape smartArtShape in smartArtShapes)
        {
            // Ensure the SmartArt drawing is up‑to‑date.
            smartArtShape.UpdateSmartArtDrawing();

            // Obtain a renderer for the shape. The GetShapeRenderer method returns a ShapeRenderer
            // instance (available in Aspose.Words 23.5 and later).
            var renderer = smartArtShape.GetShapeRenderer();

            string svgFileName = $"SmartArt_{index}.svg";

            // Render the SmartArt shape to SVG.
            renderer.Save(svgFileName, svgOptions);

            Console.WriteLine($"Saved SmartArt shape #{index} as {svgFileName}");
            index++;
        }

        Console.WriteLine("Extraction completed.");
    }
}
