using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Retrieve all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the shapes and remove those that are text boxes.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove(); // Removes the shape from its parent node.
        }

        // Save the modified document back as a DOCM file.
        doc.Save("Output.docm");
    }
}
