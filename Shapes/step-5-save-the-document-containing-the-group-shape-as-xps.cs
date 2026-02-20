using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class SaveGroupShapeAsXps
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph where the group shape will be placed.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the group shape.");

        // Create a GroupShape that belongs to the document.
        GroupShape group = new GroupShape(doc)
        {
            // Set the size of the group shape (in points).
            Width = 200,
            Height = 100,
            // Position the group shape relative to the page.
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page,
            Left = 50,
            Top = 50,
            // Optional: give the group a name.
            Name = "MyGroupShape"
        };

        // Add the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // (Optional) Add a simple shape inside the group to demonstrate content.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 80,
            Height = 40,
            Left = 10,
            Top = 10,
            FillColor = System.Drawing.Color.LightBlue
        };
        group.AppendChild(innerShape);

        // Save the document as XPS using XpsSaveOptions.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("GroupShapeDocument.xps", xpsOptions);
    }
}
