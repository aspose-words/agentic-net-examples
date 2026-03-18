using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleObjectExtractor
{
    static void Main()
    {
        // Load the Word document that contains OLE objects.
        // Replace with the actual path to your document.
        Document doc = new Document("InputDocument.docx");

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Only process shapes that contain an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Build a temporary file name using the suggested extension for the OLE object.
            string tempFileName = Guid.NewGuid().ToString() + oleFormat.SuggestedExtension;
            string tempFilePath = Path.Combine(Path.GetTempPath(), tempFileName);

            // Save the raw OLE data directly to the temporary file.
            // This uses the OleFormat.Save(string) overload as defined in the API.
            oleFormat.Save(tempFilePath);

            // The temporary file now contains the binary data of the OLE object.
            Console.WriteLine($"OLE object saved to temporary file: {tempFilePath}");
        }
    }
}
