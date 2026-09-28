using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an online YouTube video.
        // The InsertOnlineVideo method requires the video URL and the desired width and height.
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
        double width = 400;   // Width of the video placeholder in points.
        double height = 300;  // Height of the video placeholder in points.
        builder.InsertOnlineVideo(videoUrl, width, height);

        // Save the document.
        doc.Save("OnlineVideo.docx");
    }
}
