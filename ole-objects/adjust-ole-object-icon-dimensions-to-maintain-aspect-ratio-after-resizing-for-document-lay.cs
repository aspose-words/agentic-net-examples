using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class AdjustOleIconAspectRatio
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary file to be used as the OLE object source.
        string tempDir = Path.GetTempPath();
        string oleFilePath = Path.Combine(tempDir, "SampleOleObject.txt");
        File.WriteAllText(oleFilePath, "This is a sample OLE object content.");

        // Insert the OLE object as an icon. The icon will be the default one chosen by Aspose.Words.
        // Parameters: file name, isLinked (false = embed), iconFile (null = default), iconCaption (null = file name).
        Shape oleShape = builder.InsertOleObjectAsIcon(oleFilePath, false, null, null);

        // Ensure the shape's aspect ratio is locked so that resizing preserves the original proportions.
        oleShape.AspectRatioLocked = true;

        // Resize the icon to a desired width (in points). Height will adjust automatically.
        oleShape.Width = 150; // 150 points ≈ 2.08 inches.

        // Optionally, you can also move the icon to a specific location on the page.
        oleShape.WrapType = WrapType.None;
        oleShape.Left = 100;
        oleShape.Top = 100;

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdjustedOleIcon.docx");
        doc.Save(outputPath);
    }
}
