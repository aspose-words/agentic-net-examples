using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the DOCX file. Adjust if the file is located elsewhere.
        const string docPath = "input.docx";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(docPath))
        {
            Console.WriteLine($"Document file not found: {Path.GetFullPath(docPath)}");
            return;
        }

        // Load the document from the specified path.
        Document doc = new Document(docPath);

        // Retrieve all Shape nodes; OLE objects are stored as shapes with an OLE format.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and process those that contain OLE data.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.OleFormat != null)
            {
                // ProgId identifies the type of OLE object (e.g., "Excel.Sheet.12").
                string progId = shape.OleFormat.ProgId ?? "Unknown";

                // Width and Height are in points (1 point = 1/72 inch).
                double widthPoints = shape.Width;
                double heightPoints = shape.Height;

                Console.WriteLine($"OLE Object ProgId: {progId}, Size: {widthPoints}x{heightPoints} points");
            }
        }
    }
}
