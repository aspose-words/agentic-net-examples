using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // URL of the online video to embed in every document.
    private const string VideoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

    // Desired size of the video placeholder (in points).
    private const double VideoWidth = 320;   // 320 points ≈ 4.44 inches
    private const double VideoHeight = 180;  // 180 points ≈ 2.5 inches

    // Input folder containing the source Word documents.
    private static readonly string InputFolder = Path.Combine(Environment.CurrentDirectory, "InputDocs");

    // Output folder where the processed documents will be saved.
    private static readonly string OutputFolder = Path.Combine(Environment.CurrentDirectory, "OutputDocs");

    public static void Main()
    {
        // Ensure the input and output directories exist.
        Directory.CreateDirectory(InputFolder);
        Directory.CreateDirectory(OutputFolder);

        // Process each .docx file in the input folder.
        foreach (string inputFilePath in Directory.GetFiles(InputFolder, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(inputFilePath);

            // Create a DocumentBuilder for editing the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph to separate the video from existing content.
            builder.InsertParagraph();
            builder.Writeln("Embedded online video:");

            // Insert the online video using the overload that takes URL, width, and height.
            builder.InsertOnlineVideo(VideoUrl, VideoWidth, VideoHeight);

            // Determine the output file path (preserve original file name).
            string outputFilePath = Path.Combine(OutputFolder, Path.GetFileName(inputFilePath));

            // Save the modified document.
            doc.Save(outputFilePath);
        }
    }
}
