using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Get all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove any shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the textbox shape from its parent node.
                shape.Remove();
            }
        }

        // Save the modified document back to DOC format.
        doc.Save("OutputDocument.doc");
    }
}
