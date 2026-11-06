using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder containing source Word documents (relative to the executable location)
        string inputFolder = Path.GetFullPath("InputDocs");
        // Folder where processed documents will be saved
        string outputFolder = Path.GetFullPath("OutputDocs");
        // URL of the online video to embed
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify that the input directory exists; if not, inform the user and exit
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            Console.WriteLine("Create the folder and place .docx files inside before running the program.");
            return;
        }

        // Get all .docx files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.docx", SearchOption.TopDirectoryOnly);

        // If no files are found, inform the user and exit
        if (files.Length == 0)
        {
            Console.WriteLine($"No .docx files found in: {inputFolder}");
            return;
        }

        foreach (string filePath in files)
        {
            // Load the Word document
            Document doc = new Document(filePath);

            // Create a DocumentBuilder for editing
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move cursor to the end of the document
            builder.MoveToDocumentEnd();

            // Add a blank paragraph before inserting the video
            builder.Writeln();

            // Insert the online video with specified size (width, height)
            Shape videoShape = builder.InsertOnlineVideo(videoUrl, 400, 300);

            // Build output file name
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string outPath = Path.Combine(outputFolder, $"{fileName}_video.docx");

            // Save the modified document
            doc.Save(outPath);
        }

        Console.WriteLine("Processing completed successfully.");
    }
}
