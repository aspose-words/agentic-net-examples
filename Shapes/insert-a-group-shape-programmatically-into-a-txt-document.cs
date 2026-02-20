using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class InsertGroupShapeIntoTxt
{
    static void Main()
    {
        // Load an existing TXT document.
        // TxtLoadOptions are used to specify any loading preferences for plain text files.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document doc = new Document("input.txt", loadOptions);

        // Create a new group shape that will hold other shapes.
        // The constructor requires a DocumentBase (the document we are working with).
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape.
        group.Width = 200;   // Width in points.
        group.Height = 100;  // Height in points.
        group.Left = 50;     // Distance from the left margin.
        group.Top = 50;      // Distance from the top margin.

        // Define the coordinate space for child shapes inside the group.
        // CoordOrigin is the top‑left corner of the group’s internal coordinate system.
        // CoordSize defines the width and height of that internal coordinate system.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(200, 100);

        // Create a rectangle shape to place inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 20;   // Position relative to the group's coordinate space.
        rect.Top = 20;
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Insert the group shape into the document.
        // Here we add it after the first paragraph of the first section.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.ParentNode.InsertAfter(group, firstParagraph);

        // Save the modified document back to TXT format.
        // TxtSaveOptions can be used to control how the document is exported as plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("output.txt", saveOptions);
    }
}
