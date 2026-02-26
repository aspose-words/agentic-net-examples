using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsDocm
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Collect all shapes that are text boxes.
        List<Shape> textBoxShapes = new List<Shape>();
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each text box from its parent node.
        foreach (Shape textBox in textBoxShapes)
        {
            // The shape may be inside a paragraph; removing it will keep the paragraph intact.
            textBox.Remove();
        }

        // Save the modified document as a macro‑enabled DOCM file.
        doc.Save("OutputDocument.docm");
    }
}
