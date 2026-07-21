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

        // Ensure both input and output directories exist.
        // If the input folder does not exist, create it so that the program does not throw.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // URL of the online video to embed.
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Process each .docx file in the input folder.
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing document.
            Document doc = new Document(inputPath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the end of the document and add a new paragraph.
            builder.MoveToDocumentEnd();
            builder.InsertParagraph();

            // Optional descriptive text.
            builder.Writeln("Embedded video:");

            // Insert the online video with specified size and text wrapping.
            builder.InsertOnlineVideo(
                videoUrl,
                RelativeHorizontalPosition.LeftMargin, 0,
                RelativeVerticalPosition.TopMargin, 0,
                320,   // width in points
                180,   // height in points
                WrapType.Square);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            doc.Save(outputPath);
        }
    }
}
