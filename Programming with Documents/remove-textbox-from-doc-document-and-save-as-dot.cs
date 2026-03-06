using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Find all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Remove only shapes that are text boxes.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as a Word template (DOT format).
        doc.Save("output.dot", SaveFormat.Dot);
    }
}
