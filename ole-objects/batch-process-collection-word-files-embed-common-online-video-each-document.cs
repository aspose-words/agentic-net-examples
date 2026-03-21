using System;
using System.IO;
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

        // Use folders relative to the executable location
        string baseDir = AppContext.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure both directories exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no .docx files, create a simple one for demonstration
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            var demoDoc = new Document();
            var builder = new DocumentBuilder(demoDoc);
            builder.Writeln("Demo document created automatically.");
            string demoPath = Path.Combine(inputFolder, "Demo.docx");
            demoDoc.Save(demoPath);
        }

        // Retrieve all .docx files from the input folder
        string[] sourceFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string sourcePath in sourceFiles)
        {
            // Load the existing document
            Document doc = new Document(sourcePath);

            // Create a DocumentBuilder positioned at the end of the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // Insert the online video
            builder.InsertOnlineVideo(videoUrl, videoWidth, videoHeight);

            // Construct the output file path (same file name, different folder)
            string fileName = Path.GetFileName(sourcePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the modified document
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing complete. Check the 'Output' folder.");
    }
}
