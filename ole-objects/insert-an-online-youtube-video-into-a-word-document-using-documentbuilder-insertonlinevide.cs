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

        // Insert an online YouTube video. Width and height are specified in points.
        builder.InsertOnlineVideo("https://youtu.be/g1N9ke8Prmk", 360, 270);

        // Add a paragraph break after the video (optional).
        builder.Writeln();

        // Save the document to the local file system.
        doc.Save("OnlineVideo.docx");
    }
}
