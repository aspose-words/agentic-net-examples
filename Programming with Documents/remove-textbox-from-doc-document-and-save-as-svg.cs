using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.docx");

        // Find all shapes that are text boxes and remove them from the document.
        var textBoxShapes = doc.GetChildNodes(NodeType.Shape, true)
                               .Cast<Shape>()
                               .Where(s => s.ShapeType == ShapeType.TextBox)
                               .ToList();

        foreach (var shape in textBoxShapes)
            shape.Remove();

        // Configure SVG save options (no page border, fit to viewport, render text as placed glyphs).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            FitToViewPort = true,
            ShowPageBorder = false,
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the modified document as an SVG file.
        doc.Save("Output.svg", svgOptions);
    }
}
