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

        // Prepare some sample data to embed as an OLE object (a simple text file).
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE embedded text file.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object. Use the generic "Package" progId for a file container.
            // asIcon = false (display the object itself), presentation = null (default icon if needed).
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", false, null);

            // Lock the aspect ratio to preserve original proportions.
            oleShape.AspectRatioLocked = true;
        }

        // Save the document to disk.
        string outputPath = "OleObjectAspectRatio.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created successfully.");
    }
}
