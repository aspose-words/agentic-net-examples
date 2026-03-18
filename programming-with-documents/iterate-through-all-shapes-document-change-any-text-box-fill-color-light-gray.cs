using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("{inputPath}");

        // Get all shapes in the document (including those inside text boxes).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Check if the shape is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Change the fill to a solid light gray color.
                shape.Fill.Solid(Color.LightGray);
            }
        }

        // Save the modified document.
        doc.Save("{outputPath}");
    }
}
