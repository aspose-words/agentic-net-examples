using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Locate all shapes that are OLE objects.
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .Cast<Shape>()
                           .Where(s => s.ShapeType == ShapeType.OleObject)
                           .ToList();

        // Remove each OLE object shape from the document.
        foreach (var shape in oleShapes)
        {
            shape.Remove();
        }

        // Save the document after OLE objects have been deleted.
        doc.Save("output.docx");
    }
}
