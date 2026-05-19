using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://vimeo.com/52477838";

        // Insert the video as a floating shape with explicit size and position.
        // The shape will be placed relative to the left and top margins,
        // with a width of 320 points and a height of 180 points.
        builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin,   // Horizontal reference.
            0,                                      // Distance from the left margin.
            RelativeVerticalPosition.TopMargin,     // Vertical reference.
            0,                                      // Distance from the top margin.
            320,                                    // Width in points.
            180,                                    // Height in points.
            WrapType.Square);                       // Text wrapping style.

        // Save the document to the file system.
        doc.Save("OnlineVideo.docx");
    }
}
