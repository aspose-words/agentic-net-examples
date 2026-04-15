using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some data to embed as an OLE package.
        byte[] data = System.Text.Encoding.UTF8.GetBytes("Sample OLE data");
        using (MemoryStream dataStream = new MemoryStream(data))
        {
            // Insert the OLE object (not as an icon) using the generic "Package" progId.
            builder.InsertOleObject(dataStream, "Package", false, null);
        }

        // Locate the first shape that contains an OLE object.
        Shape oleShape = doc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.OleFormat != null);

        if (oleShape != null)
        {
            OleFormat oleFormat = oleShape.OleFormat;

            // Build an output file name using the suggested extension, if any.
            string outputFile = "ExtractedOle" + oleFormat.SuggestedExtension;

            // Save the OLE object's binary data to a file via a stream.
            using (FileStream fs = new FileStream(outputFile, FileMode.Create))
            {
                oleFormat.Save(fs);
            }

            Console.WriteLine($"OLE object extracted to {outputFile}");
        }
    }
}
