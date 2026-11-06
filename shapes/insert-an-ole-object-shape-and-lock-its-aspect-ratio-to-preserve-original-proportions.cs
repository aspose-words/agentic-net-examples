using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary text file to embed as an OLE object.
        string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);
        string oleFilePath = Path.Combine(tempDir, "Sample.txt");
        File.WriteAllText(oleFilePath, "This is sample text for OLE embedding.");

        // Insert the OLE object shape using a stream overload.
        Shape oleShape;
        using (FileStream oleStream = File.OpenRead(oleFilePath))
        {
            // "Package" is a generic progId that works for plain files.
            // The last parameter (icon stream) is null – Word will generate a default icon.
            oleShape = builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Lock the aspect ratio to preserve original proportions.
        oleShape.AspectRatioLocked = true;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectShape.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Clean up temporary files.
        try { File.Delete(oleFilePath); } catch { }
        try { Directory.Delete(tempDir, true); } catch { }
    }
}
