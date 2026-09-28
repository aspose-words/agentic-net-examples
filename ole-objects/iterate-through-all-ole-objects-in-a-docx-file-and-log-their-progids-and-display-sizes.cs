using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the DOCX file to process.
        const string inputPath = "input.docx";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes in the document (OLE objects are stored as Shape nodes).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each Shape node and process those that are OLE objects.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null)
            {
                // ProgId identifies the type of OLE object (e.g., Excel.Sheet.12).
                string progId = shape.OleFormat.ProgId;

                // Width and Height are measured in points.
                double width = shape.Width;
                double height = shape.Height;

                Console.WriteLine($"OLE Object ProgId: {progId}, Size: {width:F2}pt x {height:F2}pt");
            }
        }
    }
}
