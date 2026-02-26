using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class DeleteOleObjectExample
{
    static void Main()
    {
        // Load the DOCX document that contains OLE objects.
        Document doc = new Document("InputWithOle.docx");

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate over the collection and remove every shape that is an OLE object.
        // Shape.ShapeType == ShapeType.OleObject identifies OLE objects.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            if (shape.ShapeType == ShapeType.OleObject)
            {
                // Remove the shape (and thus the embedded OLE object) from its parent.
                shape.Remove();
            }
        }

        // Save the modified document.
        doc.Save("OutputWithoutOle.docx");
    }
}
