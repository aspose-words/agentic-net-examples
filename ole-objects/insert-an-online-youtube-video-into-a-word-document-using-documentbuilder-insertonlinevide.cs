using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OnlineVideoExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // URL of the YouTube video to embed.
            string videoUrl = "https://youtu.be/g1N9ke8Prmk";

            // Insert the online video with specified width and height (in points).
            // Width = 360 points, Height = 270 points.
            builder.InsertOnlineVideo(videoUrl, 360, 270);

            // Save the document to the local file system.
            doc.Save("OnlineVideo.docx");
        }
    }
}
