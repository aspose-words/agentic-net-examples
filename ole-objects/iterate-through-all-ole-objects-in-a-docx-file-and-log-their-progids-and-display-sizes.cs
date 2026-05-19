using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE object.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample OLE data");
        using (MemoryStream oleStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object as an icon. The ProgId "Package" denotes a generic OLE package.
            builder.InsertOleObject(oleStream, "Package", true, null);
        }

        // Iterate through all shapes in the document and find OLE objects.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Log the ProgId and the display size (width and height) of the OLE object.
                Console.WriteLine($"ProgId: {oleFormat.ProgId}, Size: {shape.Width} x {shape.Height}");
            }
        }

        // Save the document to the local file system.
        doc.Save("OleObjectsDemo.docx");
    }
}
