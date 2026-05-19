using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class AdjustOleIconAspectRatio
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleExample");
        Directory.CreateDirectory(tempFolder);

        // Create a simple text file that will be embedded as an OLE object.
        string oleFilePath = Path.Combine(tempFolder, "Sample.txt");
        File.WriteAllText(oleFilePath, "This is sample content for the OLE object.");

        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon. Passing null for the icon file lets Aspose.Words use a default icon.
        // The caption will be the file name.
        Shape oleShape = builder.InsertOleObjectAsIcon(oleFilePath, false, null, "Sample.txt");

        // Lock the aspect ratio so that changing one dimension automatically updates the other.
        oleShape.AspectRatioLocked = true;

        // Resize the icon width; the height will adjust to preserve the original aspect ratio.
        oleShape.Width = 100.0; // Width in points (1 point = 1/72 inch)

        // Save the document to the temporary folder.
        string outputPath = Path.Combine(tempFolder, "OleIconAdjusted.docx");
        doc.Save(outputPath);

        // The example finishes without waiting for user input.
    }
}
