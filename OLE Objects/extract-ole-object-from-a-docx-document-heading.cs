using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the DOCX document that contains OLE objects.
        // The LoadOptions class is not required here, so we use the standard constructor.
        Document doc = new Document("InputDocument.docx");

        // Iterate through all Shape nodes in the document.
        // Shapes that contain OLE objects have a non‑null OleFormat property.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if this shape actually holds an OLE object.
            if (shape.OleFormat == null)
                continue;

            OleFormat ole = shape.OleFormat;

            // Build a file name for the extracted OLE object.
            // Use the suggested file extension if it is available; otherwise default to .bin.
            string extension = string.IsNullOrEmpty(ole.SuggestedExtension) ? ".bin" : ole.SuggestedExtension;
            string fileName = $"Extracted_{Guid.NewGuid()}{extension}";

            // Save the OLE object data directly to a file.
            // This uses the OleFormat.Save(string) method provided by Aspose.Words.
            ole.Save(fileName);

            Console.WriteLine($"Extracted OLE object saved as: {fileName}");
        }
    }
}
