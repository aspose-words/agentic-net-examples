using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsPng
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Find all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a text box.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as a PNG image (first page rendered).
        doc.Save("OutputImage.png", SaveFormat.Png);
    }
}
