using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Find all shapes in the document and remove those that are text boxes.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the cleaned document as PDF. The format is inferred from the file extension.
        doc.Save("output.pdf");
    }
}
