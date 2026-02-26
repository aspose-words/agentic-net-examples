using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Find all shapes that contain OLE objects.
        List<Shape> oleShapes = new List<Shape>();
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.OleObject)
                oleShapes.Add(shape);
        }

        // Remove each OLE shape from the document.
        foreach (Shape shape in oleShapes)
            shape.Remove();

        // Save the document without the OLE objects.
        doc.Save("Output.docx");
    }
}
