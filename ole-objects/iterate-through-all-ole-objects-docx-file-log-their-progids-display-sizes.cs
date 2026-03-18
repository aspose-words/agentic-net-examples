using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Iterate over all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // If the shape contains an OLE object, its OleFormat property will be non‑null.
            OleFormat ole = shape.OleFormat;
            if (ole != null)
            {
                // Log the ProgId of the OLE object.
                Console.WriteLine($"OLE ProgId: {ole.ProgId}");

                // Log the display size of the shape (width and height in points).
                Console.WriteLine($"Display size: {shape.Width} x {shape.Height} points");
            }
        }
    }
}
