using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class AdjustOleIconAspectRatio
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary text file to embed as an OLE object.
        string tempDir = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles");
        Directory.CreateDirectory(tempDir);
        string oleFilePath = Path.Combine(tempDir, "Sample.txt");
        File.WriteAllText(oleFilePath, "This is a sample text file for OLE embedding.");

        // Insert the OLE object as an icon. No custom icon file is provided (null), so Aspose.Words uses a default one.
        // The returned Shape represents the OLE object icon.
        Shape oleShape = builder.InsertOleObjectAsIcon(oleFilePath, false, null, "Sample Text File");

        // Lock the aspect ratio to keep the icon proportions consistent when resizing.
        oleShape.AspectRatioLocked = true;

        // Desired new width for the icon (in points). Height will be adjusted to preserve the aspect ratio.
        double desiredWidth = 150.0;
        double originalWidth = oleShape.Width;
        double scaleFactor = desiredWidth / originalWidth;

        // Apply the new dimensions while maintaining the original aspect ratio.
        oleShape.Width = desiredWidth;
        oleShape.Height = oleShape.Height * scaleFactor;

        // Save the document to the output file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleIconAdjusted.docx");
        doc.Save(outputPath);
    }
}
