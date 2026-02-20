using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class AddGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a paragraph where the group shape will be placed.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is a group shape containing two rectangles:");

        // Create a GroupShape. The constructor requires a DocumentBase (the document).
        GroupShape group = new GroupShape(doc);

        // Set the size of the group shape (in points). This defines the outer bounding box.
        group.Width = 300;   // 300 points wide
        group.Height = 200;  // 200 points high

        // Set the coordinate space inside the group. This is similar to a canvas.
        // CoordOrigin is the top‑left corner of the canvas (usually 0,0).
        group.CoordOrigin = new Point(0, 0);
        // CoordSize defines the width and height of the canvas in points.
        // Here we use the same size as the outer bounds for simplicity.
        group.CoordSize = new Size(300, 200);

        // -----------------------------------------------------------------
        // Add child shapes to the group.
        // -----------------------------------------------------------------

        // First rectangle.
        Shape rect1 = new Shape(doc, ShapeType.Rectangle);
        rect1.Width = 100;
        rect1.Height = 80;
        rect1.Left = 20;   // Position within the group canvas.
        rect1.Top = 20;
        rect1.FillColor = Color.LightBlue;
        rect1.StrokeColor = Color.DarkBlue;
        group.AppendChild(rect1);

        // Second rectangle.
        Shape rect2 = new Shape(doc, ShapeType.Rectangle);
        rect2.Width = 120;
        rect2.Height = 60;
        rect2.Left = 150;
        rect2.Top = 100;
        rect2.FillColor = Color.LightCoral;
        rect2.StrokeColor = Color.DarkRed;
        group.AppendChild(rect2);

        // -----------------------------------------------------------------
        // Insert the group shape into the document.
        // -----------------------------------------------------------------
        // The group shape must be added to a paragraph (as a child node).
        Paragraph para = builder.CurrentParagraph;
        para.AppendChild(group);

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        // To ensure the group shape is saved using DrawingML (required for
        // non‑primitive shapes), set the OOXML compliance to a version that
        // supports DrawingML.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        doc.Save("GroupShapeExample.docx", saveOptions);
    }
}
