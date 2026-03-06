using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled DOCM file.
        Document doc = new Document("Input.docm");

        // Collect all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the index order.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            // Remove only shapes that are text boxes.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as a DOT template.
        doc.Save("Output.dot", SaveFormat.Dot);
    }
}
