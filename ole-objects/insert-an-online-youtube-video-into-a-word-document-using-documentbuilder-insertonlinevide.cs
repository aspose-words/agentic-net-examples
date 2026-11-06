using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an online YouTube video.
        // Parameters: video URL, width (points), height (points).
        builder.InsertOnlineVideo("https://www.youtube.com/watch?v=dQw4w9WgXcQ", 400, 300);

        // Save the document to disk.
        doc.Save("OnlineVideo.docx");
    }
}
