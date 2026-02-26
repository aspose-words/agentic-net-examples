using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the DOCX document that contains OLE objects.
        Document doc = new Document("InputDocument.docx");

        // Get all shape nodes in the document (OLE objects are stored as shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and process only those that are OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType != ShapeType.OleObject)
                continue; // Skip non‑OLE shapes.

            // Access the OleFormat of the shape.
            OleFormat oleFormat = shape.OleFormat;

            // Determine a suitable file name using the suggested extension.
            string outputFileName = $"ExtractedObject_{shape.GetHashCode()}{oleFormat.SuggestedExtension}";

            // Save the embedded OLE object directly to a file.
            oleFormat.Save(outputFileName);

            // Optional: also demonstrate saving via a stream.
            // using (FileStream fs = new FileStream("Stream_" + outputFileName, FileMode.Create))
            // {
            //     oleFormat.Save(fs);
            // }

            Console.WriteLine($"OLE object saved to: {outputFileName}");
        }
    }
}
