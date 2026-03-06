using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class AddGroupShapeExample
{
    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 3. Insert two individual shapes that we will later group.
        //    These shapes must be DML (DrawingML) if we intend to save with ISO compliance.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // 4. Group the shapes using DocumentBuilder.InsertGroupShape.
        //    The method automatically calculates the group’s position and size.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // 5. (Optional) Adjust the group’s coordinate system if you need precise control.
        //    CoordSize defines the internal coordinate space; CoordOrigin sets its origin.
        group.CoordSize = new Size(500, 500);          // internal grid 500x500 points
        group.CoordOrigin = new Point(-250, -250);    // origin shifted to centre

        // 6. Verify that the created node is indeed a group shape.
        if (!group.IsGroup)
            throw new InvalidOperationException("The node is not a group shape.");

        // 7. Set OOXML compliance if you plan to save the document in a strict format.
        //    DML shapes are required for ISO/IEC 29500 compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // 8. Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx", saveOptions);
    }
}
