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

        // Insert an online YouTube video with specified width and height (in points).
        builder.InsertOnlineVideo("https://youtu.be/g1N9ke8Prmk", 360, 270);

        // Save the document to a file in the current directory.
        doc.Save("OnlineVideo.docx");
    }
}
