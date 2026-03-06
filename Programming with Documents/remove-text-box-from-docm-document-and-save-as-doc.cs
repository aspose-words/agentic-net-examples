using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled DOCM file.
        Document doc = new Document("Input.docm");

        // Get all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removing a shape does not affect the loop index.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Remove only text box shapes.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as a legacy DOC (macro‑free) file.
        doc.Save("Output.doc", SaveFormat.Doc);
    }
}
