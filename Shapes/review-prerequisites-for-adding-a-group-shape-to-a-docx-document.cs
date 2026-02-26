using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document – the first prerequisite.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure OOXML compliance to allow DML shapes (required for non‑primitive shapes).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // Insert two floating shapes that can be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Ensure both shapes are top‑level (not already inside another group).
        if (!shape1.IsTopLevel || !shape2.IsTopLevel)
            throw new InvalidOperationException("Shapes must be top‑level before grouping.");

        // Group the shapes using the built‑in InsertGroupShape method.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Verify that the created node is indeed a group shape.
        if (!group.IsGroup)
            throw new InvalidOperationException("Failed to create a group shape.");

        // Optional: adjust group bounds and coordinate system if specific sizing is needed.
        group.Bounds = new RectangleF(10, 10, 300, 300);
        group.CoordSize = new Size(500, 500);
        group.CoordOrigin = new Point(-250, -250);

        // Save the document using the configured compliance options.
        doc.Save("GroupShapePrerequisites.docx", saveOptions);
    }
}
