using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline shape (DiagonalCornersRounded) with a size of 50x50 points.
        Shape inlineShape = builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);
        inlineShape.FillColor = Color.LightBlue; // Optional formatting.

        // Insert a floating shape (TopCornersRounded) positioned 100 points from the page edges.
        Shape floatingShape = builder.InsertShape(
            ShapeType.TopCornersRounded,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            50, 50,
            WrapType.None);
        floatingShape.FillColor = Color.LightGreen; // Optional formatting.

        // Create two more shapes to demonstrate grouping.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);
        group.Left = 300; // Position the group on the page.
        group.Top = 100;

        // Save the document using OOXML compliance that supports DML shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("ShapesInsertion.docx", saveOptions);
    }
}
