using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class DeleteOleObjectExample
{
    static void Main()
    {
        // Load the DOCX document (uses the provided Document constructor)
        Document doc = new Document("Input.docx");

        // Collect all Shape nodes that represent OLE objects
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .Cast<Shape>()
                           .Where(s => s.ShapeType == ShapeType.OleObject)
                           .ToList();

        // Remove each OLE shape from the document tree
        foreach (var shape in oleShapes)
        {
            shape.Remove();
        }

        // Save the modified document (uses the provided Save method)
        doc.Save("Output.docx");
    }
}
