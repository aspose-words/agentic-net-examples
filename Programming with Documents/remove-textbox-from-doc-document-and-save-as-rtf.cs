using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsRtf
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Find all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Check if the shape is a TextBox.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the textbox shape from its parent.
                shape.Remove();
            }
        }

        // Save the modified document as RTF.
        doc.Save("OutputDocument.rtf");
    }
}
