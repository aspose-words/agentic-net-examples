using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder that contains the source Word documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        // Folder where the processed documents will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // URL of the online video to embed in each document.
        const string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Process every .docx file in the input folder.
        foreach (string inputFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(inputFile);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the video.
            builder.Writeln("Embedded online video:");

            // Insert the online video with a fixed size (width: 320 points, height: 180 points).
            // This uses the InsertOnlineVideo(string, double, double) overload.
            builder.InsertOnlineVideo(videoUrl, 320, 180);

            // Determine the output file path (same file name, different folder).
            string outputFile = Path.Combine(outputFolder, Path.GetFileName(inputFile));

            // Save the modified document.
            doc.Save(outputFile);
        }
    }
}
