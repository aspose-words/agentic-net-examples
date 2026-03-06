using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document (uses the provided Document(string) constructor)
        Document doc = new Document("Input.docx");

        // Retrieve all shapes that are OLE objects
        var oleShapes = doc.GetChildNodes(NodeType.Shape, true)
                           .OfType<Shape>()
                           .Where(s => s.ShapeType == ShapeType.OleObject)
                           .ToList();

        // Delete each OLE object shape from the document
        foreach (var shape in oleShapes)
        {
            shape.Remove();
        }

        // Save the modified document (uses the provided Document.Save(string) method)
        doc.Save("Output.docx");
    }
}
