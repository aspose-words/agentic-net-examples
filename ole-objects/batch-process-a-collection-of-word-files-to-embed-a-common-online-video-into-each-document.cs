using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder containing the source Word documents.
        string inputFolder = "InputDocs";

        // Folder where the processed documents will be saved.
        string outputFolder = "OutputDocs";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // URL of the online video to embed (YouTube or Vimeo supported).
        string videoUrl = "https://youtu.be/g1N9ke8Prmk";

        // Process each .docx file in the input folder.
        foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(sourcePath);

            // Create a DocumentBuilder for editing.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document.
            builder.MoveToDocumentEnd();

            // Optional: add a paragraph before the video.
            builder.Writeln();
            builder.Writeln("Embedded online video:");

            // Insert the online video with a width of 320 points and height of 180 points.
            builder.InsertOnlineVideo(videoUrl, 320, 180);

            // Determine the output file path (same file name, different folder).
            string fileName = Path.GetFileName(sourcePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
