using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path to the DOCX file to process.
        const string inputPath = "input.docx";

        // Load the document.
        Document doc;
        try
        {
            doc = new Document(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading document: {ex.Message}");
            return;
        }

        // Retrieve all Shape nodes (OLE objects are stored as shapes).
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int oleIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Check if the shape contains an OLE object.
            if (shape.OleFormat != null)
            {
                oleIndex++;
                string progId = shape.OleFormat.ProgId ?? "N/A";

                // Width and Height are in points (1 point = 1/72 inch).
                double widthInches = shape.Width / 72.0;
                double heightInches = shape.Height / 72.0;

                Console.WriteLine($"OLE Object {oleIndex}: ProgId = {progId}, Size = {widthInches:F2}in x {heightInches:F2}in");
            }
        }

        if (oleIndex == 0)
        {
            Console.WriteLine("No OLE objects found in the document.");
        }
    }
}
