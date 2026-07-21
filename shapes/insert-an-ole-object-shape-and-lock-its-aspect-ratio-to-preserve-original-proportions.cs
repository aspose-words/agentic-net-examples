using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a simple text file content to embed as an OLE object.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("This is a sample OLE object content.");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Ensure the stream is positioned at the beginning.
            oleStream.Position = 0;

            // Insert the OLE object. Use the generic "Package" progId for a plain file.
            // asIcon = false to embed the object directly, presentation = null to use default icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Lock the aspect ratio to preserve original proportions.
            oleShape.AspectRatioLocked = true;
        }

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");
    }
}
