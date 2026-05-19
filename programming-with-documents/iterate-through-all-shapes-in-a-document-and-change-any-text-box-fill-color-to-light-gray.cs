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

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.WrapType = WrapType.None;
        textBox.FillColor = Color.LightBlue; // Initial fill color.

        // Insert another shape (rectangle) to demonstrate that only text boxes are affected.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        rectangle.FillColor = Color.LightCoral;

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Check if the shape is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Change the fill color of the text box to light gray.
                shape.FillColor = Color.LightGray;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
