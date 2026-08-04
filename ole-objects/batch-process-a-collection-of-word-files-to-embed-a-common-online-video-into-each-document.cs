using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder containing the source Word documents.
        // Use a path relative to the current working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        // Folder where the modified documents will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure both directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // URL of the online video to embed in each document.
        const string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Process every .docx file in the input folder.
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(inputPath);

            // Create a builder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document.
            builder.MoveToDocumentEnd();

            // Optional: add a paragraph break before the video.
            builder.Writeln();

            // Insert the online video with a fixed size (width x height in points).
            // This uses the InsertOnlineVideo(string, double, double) overload.
            builder.InsertOnlineVideo(videoUrl, 320, 180);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            doc.Save(outputPath);
        }

        // Inform the user if no documents were found.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            Console.WriteLine($"No .docx files were found in '{inputFolder}'. Place documents there and rerun the program.");
        }
        else
        {
            Console.WriteLine("Processing completed successfully.");
        }
    }
}
