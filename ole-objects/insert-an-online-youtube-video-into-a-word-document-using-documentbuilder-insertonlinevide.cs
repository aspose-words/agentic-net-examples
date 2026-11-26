using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an online YouTube video.
        // Width and height are specified in points (1 point = 1/72 inch).
        string youtubeUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
        double width = 400;   // Approx. 5.55 inches
        double height = 300;  // Approx. 4.17 inches
        builder.InsertOnlineVideo(youtubeUrl, width, height);

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
