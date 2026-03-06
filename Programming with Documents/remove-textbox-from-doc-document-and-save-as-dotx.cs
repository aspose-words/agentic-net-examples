using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Find all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as a DOTX template.
        doc.Save("OutputTemplate.dotx", SaveFormat.Dotx);
    }
}
