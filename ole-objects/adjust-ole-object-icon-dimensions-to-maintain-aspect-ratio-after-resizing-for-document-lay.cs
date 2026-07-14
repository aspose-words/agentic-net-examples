using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleOleObject.txt");
        File.WriteAllText(tempFilePath, "This is a sample OLE object content.");

        // Initialize a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon. No custom icon file is provided (null),
        // so Aspose.Words will use a predefined icon based on the file type.
        // The caption for the icon is set to the file name.
        Shape oleShape = builder.InsertOleObjectAsIcon(
            tempFilePath,          // fileName
            false,                 // isLinked
            null,                  // iconFile (use default)
            "Sample OLE Object");  // iconCaption

        // Lock the aspect ratio so that resizing the width automatically adjusts the height.
        oleShape.AspectRatioLocked = true;

        // Resize the icon to a desired width (in points). Height will scale proportionally.
        oleShape.Width = 150; // 150 points ≈ 2.08 inches.

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleIconAspectRatio.docx");
        doc.Save(outputPath);
    }
}
