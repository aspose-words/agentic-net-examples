using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsSvg
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find all shapes that are text boxes and remove them from the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true)
                                   .OfType<Shape>()
                                   .Where(s => s.ShapeType == ShapeType.TextBox))
        {
            shape.Remove();
        }

        // Configure SVG save options (optional settings).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            FitToViewPort = true,                     // Make SVG fill the viewport.
            ShowPageBorder = false,                   // Do not draw a page border.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs // Render text as curves.
        };

        // Save the modified document as an SVG file.
        doc.Save("OutputDocument.svg", svgOptions);
    }
}
