using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Collect all shapes that are text boxes.
        NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);
        List<Shape> textBoxShapes = new List<Shape>();
        foreach (Shape shape in allShapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each text box from its parent.
        foreach (Shape tb in textBoxShapes)
        {
            tb.Remove();
        }

        // Save the modified document as Markdown.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
