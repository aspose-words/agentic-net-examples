using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class AdjustOleIconAspectRatio
{
    static void Main()
    {
        // Create a temporary file to embed as an OLE object.
        string oleFilePath = Path.Combine(Path.GetTempPath(), "Sample.txt");
        File.WriteAllText(oleFilePath, "This is a sample embedded text file.");

        // No custom icon file – let Aspose.Words use the default icon.
        string? iconFilePath = null;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon. The method returns the Shape that represents the icon.
        Shape oleIconShape = builder.InsertOleObjectAsIcon(
            oleFilePath,          // fileName
            false,                // isLinked – embed the file
            iconFilePath,         // iconFile – default icon
            "Sample Text File");  // iconCaption

        // Retrieve the original dimensions of the icon image (in points).
        ImageSize originalSize = oleIconShape.ImageData.ImageSize;

        // Desired width for the icon in the document (points). 1 point = 1/72 inch.
        double desiredWidth = 72.0; // 1 inch wide

        // Lock the aspect ratio so that changing one dimension automatically updates the other.
        oleIconShape.AspectRatioLocked = true;

        // Apply the new width; height will be adjusted automatically to preserve the ratio.
        oleIconShape.Width = desiredWidth;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleIconAdjusted.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
