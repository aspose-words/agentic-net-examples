using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class FitWidthExample
{
    static void Main()
    {
        // Create a temporary image file (a tiny red dot) if it doesn't already exist.
        string tempImagePath = Path.Combine(Path.GetTempPath(), "sample.png");
        if (!File.Exists(tempImagePath))
        {
            // PNG data for a 1x1 red pixel.
            byte[] pngData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
            File.WriteAllBytes(tempImagePath, pngData);
        }

        // Define the output document path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FitWidthDocument.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image from the temporary file.
        Shape imageShape = builder.InsertImage(tempImagePath);

        // Define the maximum width (in points). 1 point = 1/72 inch.
        double maxWidthPoints = 200.0; // Adjust as needed.

        // If the image is wider than the maximum, limit its width.
        if (imageShape.Width > maxWidthPoints)
        {
            imageShape.Width = maxWidthPoints;
            imageShape.ImageData.FitImageToShape();
        }

        // Save the document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
