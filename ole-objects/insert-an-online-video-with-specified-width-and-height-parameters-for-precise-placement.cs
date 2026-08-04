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

        // Insert the online video with explicit positioning and size.
        // Parameters:
        // - videoUrl: URL of the video.
        // - RelativeHorizontalPosition.LeftMargin: distance measured from the left margin.
        // - left: 0 points from the left margin.
        // - RelativeVerticalPosition.TopMargin: distance measured from the top margin.
        // - top: 0 points from the top margin.
        // - width: 320 points (approx. 4.44 inches).
        // - height: 180 points (approx. 2.5 inches).
        // - WrapType.Square: text wraps around the video shape.
        builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin, 0,
            RelativeVerticalPosition.TopMargin, 0,
            320, 180,
            WrapType.Square);

        // Define the output path (in the current working directory).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OnlineVideo.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
