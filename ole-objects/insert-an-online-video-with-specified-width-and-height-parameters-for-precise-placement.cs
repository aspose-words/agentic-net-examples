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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the online video to embed.
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Insert the video as a floating shape.
        // Position: relative to the left and top margins, offset (0,0).
        // Size: 320 points wide by 180 points high (16:9 aspect ratio).
        // Wrap type: square text wrapping.
        builder.InsertOnlineVideo(
            videoUrl,
            RelativeHorizontalPosition.LeftMargin,
            0,
            RelativeVerticalPosition.TopMargin,
            0,
            320,
            180,
            WrapType.Square);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "OnlineVideo.docx");
        doc.Save(outputPath);
    }
}
