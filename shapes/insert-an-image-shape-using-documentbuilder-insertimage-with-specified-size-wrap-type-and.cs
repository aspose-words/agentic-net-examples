using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageShapeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a folder for the generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(outputDir);

            // Create a simple 1x1 PNG image from a base‑64 string (no System.Drawing dependency).
            string imagePath = Path.Combine(outputDir, "sample.png");
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            File.WriteAllBytes(imagePath, pngBytes);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the image as a floating shape with custom size, position and wrap type.
            // Parameters: file name, horizontal position reference, left, vertical position reference, top,
            // width, height, wrap type.
            Shape imageShape = builder.InsertImage(
                imagePath,
                RelativeHorizontalPosition.Margin, 100,   // 100 points from left margin
                RelativeVerticalPosition.Margin, 100,     // 100 points from top margin
                200,                                      // width in points
                200,                                      // height in points
                WrapType.Square);                         // text will wrap around the image

            // Ensure the image appears in front of the text.
            imageShape.BehindText = false;

            // Save the document.
            string docPath = Path.Combine(outputDir, "ImageShape.docx");
            doc.Save(docPath);

            // Validate that the output file was created.
            if (!File.Exists(docPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Optional clean‑up of the temporary image file.
            // File.Delete(imagePath);
        }
    }
}
