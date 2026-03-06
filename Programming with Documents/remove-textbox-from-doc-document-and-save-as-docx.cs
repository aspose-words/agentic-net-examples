using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("Input.doc");

        // Find all shapes that are text boxes.
        NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);
        List<Shape> textBoxShapes = new List<Shape>();

        foreach (Shape shape in allShapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each text box from the document.
        foreach (Shape textBox in textBoxShapes)
            textBox.Remove();

        // Save the modified document as DOCX.
        doc.Save("Output.docx");
    }
}
