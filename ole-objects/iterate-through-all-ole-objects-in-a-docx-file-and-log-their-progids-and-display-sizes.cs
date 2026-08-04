using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the path to the DOCX file.
        // Use the first command‑line argument if supplied; otherwise fall back to "input.docx".
        string docPath = args.Length > 0 ? args[0] : "input.docx";

        // Load the document if it exists; otherwise create an empty document.
        Document doc = File.Exists(docPath) ? new Document(docPath) : new Document();

        // Get all Shape nodes (including those in headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and output OLE information when present.
        foreach (Shape shape in shapes)
        {
            OleFormat ole = shape.OleFormat;
            if (ole != null)
            {
                Console.WriteLine($"OLE ProgId: {ole.ProgId}");
                Console.WriteLine($"Display Size: {shape.Width}pt (W) x {shape.Height}pt (H)");
                Console.WriteLine();
            }
        }
    }
}
