using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Find and remove every shape that is a TextBox.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the resulting document as a JPEG image.
        // This uses the Save(string, SaveFormat) overload, which renders the first page to JPEG.
        doc.Save("Output.jpg", SaveFormat.Jpeg);
    }
}
