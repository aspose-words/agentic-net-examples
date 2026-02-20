using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape directly (StartGroupShape/EndGroupShape are not available in this version).
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape (optional).
        group.Width = 200;   // Width in points.
        group.Height = 100;  // Height in points.
        group.Left = 100;    // Distance from the left edge of the page.
        group.Top = 100;     // Distance from the top edge of the page.

        // Create a rectangle shape to place inside the group.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 100;
        rectangle.Height = 50;
        rectangle.Left = 0;   // Position relative to the group's coordinate space.
        rectangle.Top = 0;

        // Add the rectangle to the group shape.
        group.AppendChild(rectangle);

        // Insert the group shape into the document.
        // You can either insert it at the current builder position or append it to a paragraph.
        builder.InsertNode(group);
        // Alternatively: builder.CurrentParagraph.AppendChild(group);

        // Save the document as a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
