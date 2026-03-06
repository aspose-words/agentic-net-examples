using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Collect all shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as MHTML.
        doc.Save("output.mhtml", SaveFormat.Mhtml);
    }
}
