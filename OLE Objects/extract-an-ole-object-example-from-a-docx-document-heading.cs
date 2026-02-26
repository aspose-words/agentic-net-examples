using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the DOCX document that contains OLE objects.
        Document doc = new Document("Input.docx");

        // Get all shape nodes in the document (OLE objects are stored as shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and process only those that are OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            // Check that the shape actually contains an OLE object.
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null)
            {
                OleFormat oleFormat = shape.OleFormat;

                // Determine a suitable file extension for the extracted object.
                // If the OLE object does not provide a suggestion, default to .bin.
                string extension = oleFormat.SuggestedExtension;
                if (string.IsNullOrEmpty(extension))
                    extension = ".bin";

                // Build the output file name.
                string outputFile = $"ExtractedOle_{oleIndex}{extension}";

                // Save the embedded OLE data to the file system.
                oleFormat.Save(outputFile);

                Console.WriteLine($"Saved OLE object #{oleIndex} to \"{outputFile}\".");
                oleIndex++;
            }
        }

        // Optionally, inform if no OLE objects were found.
        if (oleIndex == 0)
            Console.WriteLine("No OLE objects were found in the document.");
    }
}
