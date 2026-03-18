using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

class BatchVideoEmbed
{
    static void Main()
    {
        // URL of the online video to embed in every document
        const string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Desired size of the video placeholder (in points)
        const double videoWidth = 320;   // 320 points ≈ 4.44 inches
        const double videoHeight = 180;  // 180 points ≈ 2.5 inches

        // Folder that contains the source Word documents
        const string inputFolder = @"C:\Docs\Input";

        // Folder where the processed documents will be saved
        const string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all .docx files from the input folder
        string[] sourceFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string sourcePath in sourceFiles)
        {
            // Load the existing document (load rule)
            Document doc = new Document(sourcePath);

            // Create a DocumentBuilder positioned at the end of the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // Insert the online video using the InsertOnlineVideo overload that takes URL, width, and height
            builder.InsertOnlineVideo(videoUrl, videoWidth, videoHeight);

            // Construct the output file path (same file name, different folder)
            string fileName = Path.GetFileName(sourcePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the modified document (save rule)
            doc.Save(outputPath);
        }
    }
}
