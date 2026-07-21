using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape (the target shape type).
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.None;
        textBox.FillColor = Color.AliceBlue; // Initial fill color.
        builder.InsertNode(textBox);

        // Insert another shape (e.g., a rectangle) to demonstrate that only text boxes are changed.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        rectangle.FillColor = Color.Coral;

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Check if the shape is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Change the fill color to light gray.
                shape.FillColor = Color.LightGray;
                // Alternatively, you could use: shape.Fill.Solid(Color.LightGray);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
