using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the DOCM document from the file system.
        Document doc = new Document("Input.docm");

        // Retrieve all Shape nodes in the document (OLE objects are stored in shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int oleIndex = 0;

        // Iterate through each shape and process those that contain OLE data.
        foreach (Shape shape in shapeNodes)
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Skip shapes that are not OLE objects.

            // Construct a file name for the extracted OLE object using its suggested extension.
            string outputFileName = $"OleObject_{oleIndex}{oleFormat.SuggestedExtension}";

            // Save the embedded OLE object directly to a file.
            oleFormat.Save(outputFileName);

            oleIndex++;
        }
    }
}
