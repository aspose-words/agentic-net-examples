using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Insert the online video with precise placement.
        // Position: left margin, top margin, no offset.
        // Size: 320 points wide by 180 points high (16:9 aspect ratio).
        // Wrap type: square (text wraps around the video shape).
        builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin, 0,
            RelativeVerticalPosition.TopMargin, 0,
            320, 180,
            WrapType.Square);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OnlineVideo.docx");

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
