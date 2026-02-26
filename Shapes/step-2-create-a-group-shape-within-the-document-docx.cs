using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Create a GroupShape instance and add the previously created shapes as its children.
        GroupShape group = new GroupShape(doc);
        group.AppendChild(shape1);
        group.AppendChild(shape2);

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the document in DOCX format.
        doc.Save("GroupShape.docx", SaveFormat.Docx);
    }
}
