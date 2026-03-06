using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class NestedShapesExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the outermost group shape at the current cursor position.
        GroupShape outerGroup = builder.InsertGroupShape();

        // Set the size and position of the outer group shape.
        outerGroup.Bounds = new RectangleF(100, 100, 200, 200); // left, top, width, height

        // -------------------------------------------------
        // Create the middle shape – another GroupShape.
        // -------------------------------------------------
        GroupShape middleGroup = new GroupShape(doc);
        middleGroup.Bounds = new RectangleF(20, 20, 150, 150); // relative to the outer group

        // -------------------------------------------------
        // Create the innermost shape – a simple rectangle.
        // -------------------------------------------------
        Shape innermostShape = new Shape(doc, ShapeType.Rectangle);
        innermostShape.Width = 80;
        innermostShape.Height = 60;
        innermostShape.Left = 30;   // position inside the middle group
        innermostShape.Top = 30;
        innermostShape.Stroke.Color = Color.Blue; // visual styling

        // Add the innermost shape to the middle group.
        middleGroup.AppendChild(innermostShape);

        // Add the middle group to the outermost group.
        outerGroup.AppendChild(middleGroup);

        // Save the document containing the nested shapes.
        doc.Save("NestedShapes.docx");
    }
}
