using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class AdjustOleIconAspectRatio
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object.
        string oleFilePath = @"C:\Data\Sample.xlsx";

        // Path to a custom 32x32 (or smaller) icon file.
        string iconFilePath = @"C:\Data\CustomIcon.ico";

        // Insert the OLE object as an icon. The method returns the Shape that represents the icon.
        Shape oleIconShape = builder.InsertOleObjectAsIcon(
            oleFilePath,          // fileName
            false,                // isLinked – embed the file
            iconFilePath,         // iconFile – custom icon
            "Sample Excel File"); // iconCaption

        // Retrieve the original dimensions of the icon image (in points).
        ImageSize originalSize = oleIconShape.ImageData.ImageSize;

        // Desired width for the icon in the document (points). 1 point = 1/72 inch.
        double desiredWidth = 72.0; // 1 inch wide

        // Calculate the scaling factor based on the original width.
        double scaleFactor = desiredWidth / originalSize.WidthPoints;

        // Lock the aspect ratio so that changing one dimension automatically updates the other.
        oleIconShape.AspectRatioLocked = true;

        // Apply the new width; height will be adjusted automatically to preserve the ratio.
        oleIconShape.Width = desiredWidth;

        // (Optional) If you prefer to set height directly, you could compute it:
        // oleIconShape.Height = originalSize.HeightPoints * scaleFactor;

        // Save the document.
        doc.Save(@"C:\Output\OleIconAdjusted.docx");
    }
}
