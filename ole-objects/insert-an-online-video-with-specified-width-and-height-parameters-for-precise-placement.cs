using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertOnlineVideoExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://vimeo.com/52477838";

        // Insert the online video as a floating shape.
        // Position: relative to the left and top margins, at (0,0) offset.
        // Size: 320x180 points (16:9 aspect ratio).
        // Wrap type: Square (text wraps around the video shape).
        builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin, 0,
            RelativeVerticalPosition.TopMargin, 0,
            320, 180,
            WrapType.Square);

        // Define the output file path (saved in the current directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OnlineVideo.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
