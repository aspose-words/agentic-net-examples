using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // 1x1 pixel PNG image data (transparent)
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbV8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        using (var imageStream = new MemoryStream(imageBytes))
        {
            // Insert the image from the stream as an inline shape.
            Shape shape = builder.InsertImage(imageStream);

            // Apply the fitSizeLim behavior: fit the image data to the shape's frame.
            shape.ImageData.FitImageToShape();
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
