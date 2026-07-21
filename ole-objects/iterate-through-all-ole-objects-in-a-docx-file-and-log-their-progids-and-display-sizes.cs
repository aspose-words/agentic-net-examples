using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Resolve the path to the DOCX file relative to the executable's directory.
        // Adjust the file name as needed; the file must exist at this location.
        string inputFileName = "Sample.docx";
        string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, inputFileName);

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        // Load the document from the specified file.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and check if it contains an OLE object.
        foreach (Shape shape in shapeNodes)
        {
            OleFormat ole = shape.OleFormat;
            if (ole != null)
            {
                // ProgId of the OLE object (e.g., "Excel.Sheet.12").
                string progId = ole.ProgId;

                // Display size of the OLE object in points (1 point = 1/72 inch).
                double widthPoints = shape.Width;
                double heightPoints = shape.Height;

                Console.WriteLine($"OLE Object ProgId: {progId}, Size: {widthPoints}pt x {heightPoints}pt");
            }
        }
    }
}
