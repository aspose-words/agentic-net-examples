using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file that will be embedded as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleText.txt");
        File.WriteAllText(tempFilePath, "This is sample text for OLE embedding.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Embed the temporary text file as an OLE Package object.
        using (FileStream stream = File.OpenRead(tempFilePath))
        {
            // progId "Package" indicates a generic OLE package.
            // asIcon = false (display content), presentation = null (default icon if needed).
            builder.InsertOleObject(stream, "Package", false, null);
        }

        // Iterate over all shapes in the document to find OLE objects.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null)
            {
                // Retrieve the raw binary data of the OLE object.
                byte[] oleRawData = oleFormat.GetRawData();

                // Example custom processing: output the size of the raw data.
                Console.WriteLine($"Found OLE object. Raw data length: {oleRawData.Length} bytes.");
            }
        }

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }
    }
}
