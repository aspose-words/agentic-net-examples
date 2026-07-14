using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        string docPath = "OleDocument.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert an OLE object.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some sample data to embed as an OLE package.
        string sampleText = "This is sample text stored inside an OLE package.";
        byte[] sampleBytes = System.Text.Encoding.UTF8.GetBytes(sampleText);

        // Insert the OLE object (as a package) into the document.
        using (MemoryStream stream = new MemoryStream(sampleBytes))
        {
            // Insert as an icon (true) with no custom presentation image (null).
            // ProgId "Package" indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Optionally set a size for the displayed icon.
            oleShape.Width = 100;   // points
            oleShape.Height = 100;  // points
        }

        // Save the document containing the OLE object.
        doc.Save(docPath);

        // Load the saved document.
        Document loadedDoc = new Document(docPath);

        // Iterate through all shapes in the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Log the ProgId of the OLE object.
                Console.WriteLine($"OLE ProgId: {oleFormat.ProgId}");

                // Log the display size of the OLE object (width x height in points).
                Console.WriteLine($"Display size: {shape.Width} x {shape.Height} points");
            }
        }
    }
}
