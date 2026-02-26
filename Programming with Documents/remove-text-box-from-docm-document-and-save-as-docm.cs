using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCM file.
        Document doc = new Document("Input.docm");

        // Collect all shapes that are text boxes.
        NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);
        List<Shape> textBoxShapes = new List<Shape>();
        foreach (Shape shape in allShapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each text box from the document.
        foreach (Shape tb in textBoxShapes)
            tb.Remove();

        // Save the modified document back as DOCM.
        doc.Save("Output.docm");
    }
}
