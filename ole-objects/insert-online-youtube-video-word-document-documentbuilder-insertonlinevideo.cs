using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOnlineVideoExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – this provides the cursor for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the YouTube video to embed.
        string videoUrl = "https://youtu.be/g1N9ke8Prmk";

        // Insert the online video. Width and height are specified in points (1 point = 1/72 inch).
        // The overload InsertOnlineVideo(string, double, double) is used as defined in the API.
        builder.InsertOnlineVideo(videoUrl, 360, 270);

        // Save the document to the file system.
        doc.Save("OnlineVideo.docx");
    }
}
