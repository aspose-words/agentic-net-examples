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

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape instance attached to the document.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape (in points).
        group.Bounds = new RectangleF(0, 0, 200, 200);
        // Make the group shape floating (not inline) and disable text wrapping.
        group.WrapType = WrapType.None;

        // Create a rectangle shape that will be a child of the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;          // Width in points.
        rect.Height = 50;          // Height in points.
        rect.Left = 10;            // Position relative to the group's top‑left corner.
        rect.Top = 10;
        rect.FillColor = Color.LightBlue; // Simple fill for visibility.

        // Add the rectangle shape to the group.
        group.AppendChild(rect);

        // Insert the group shape into the document.
        // Here we add it to the current paragraph of the builder.
        builder.CurrentParagraph.AppendChild(group);

        // Save the document as DOCX with a specific OOXML compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        doc.Save("GroupShape.docx", saveOptions);
    }
}
