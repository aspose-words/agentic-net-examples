using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://vimeo.com/52477838";

        // Desired size of the video placeholder (in points).
        double width = 320;   // width of the shape
        double height = 180;  // height of the shape

        // Insert the online video as a floating shape.
        // The shape is positioned relative to the left and top margins with zero offset.
        Shape videoShape = builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin, 0,
            RelativeVerticalPosition.TopMargin, 0,
            width, height,
            WrapType.Square);

        // Additional formatting can be applied to the returned Shape if needed.
        videoShape.WrapType = WrapType.Square;

        // Save the document to disk.
        doc.Save("InsertOnlineVideo.docx");
    }
}
