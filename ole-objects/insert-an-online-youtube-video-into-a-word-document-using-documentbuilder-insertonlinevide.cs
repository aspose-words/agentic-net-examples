using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the YouTube video to embed.
        string videoUrl = "https://youtu.be/dQw4w9WgXcQ";

        // Insert the online video with a size of 320x180 points (16:9 aspect ratio).
        builder.InsertOnlineVideo(videoUrl, 320, 180);

        // Save the document to the local file system.
        doc.Save("OnlineVideo.docx");
    }
}
