using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxExample
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("InputDocument.doc");

        // Remove all shapes that are text boxes.
        // Get a live collection of all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the index of remaining items.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save("OutputDocument.docx", SaveFormat.Docx);
    }
}
