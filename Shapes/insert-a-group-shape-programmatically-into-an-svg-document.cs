using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoSvg
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape.
        group.Width = 300;   // points
        group.Height = 200;  // points
        group.Left = 50;     // points from the left margin
        group.Top = 50;      // points from the top margin

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 80;
        rect.Left = 10;   // position relative to the group
        rect.Top = 10;
        rect.Fill.ForeColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;

        // Create an ellipse shape to add to the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 80;
        ellipse.Left = 150;
        ellipse.Top = 50;
        ellipse.Fill.ForeColor = System.Drawing.Color.LightCoral;
        ellipse.Stroke.Color = System.Drawing.Color.Maroon;

        // Add the child shapes to the group.
        group.AppendChild(rect);
        group.AppendChild(ellipse);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Configure SVG save options (optional: render text as placed glyphs).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Ensure the SVG fills the viewport.
            FitToViewPort = true,
            ShowPageBorder = false
        };

        // Save the document as an SVG file.
        doc.Save("GroupShapeOutput.svg", svgOptions);
    }
}
