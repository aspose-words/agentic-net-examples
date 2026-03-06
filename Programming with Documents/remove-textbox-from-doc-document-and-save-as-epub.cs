using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Find all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as EPUB. The format is inferred from the file extension.
        doc.Save("output.epub");
    }
}
