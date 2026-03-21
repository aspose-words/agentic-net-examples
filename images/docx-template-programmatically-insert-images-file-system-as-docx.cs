using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageInsertion
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document (acts as the template).
            Document doc = new Document();

            // Create a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A 1x1 pixel PNG image (transparent) encoded in Base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            // Insert the image three times to mimic multiple images.
            for (int i = 0; i < 3; i++)
            {
                using (MemoryStream ms = new MemoryStream(pngBytes))
                {
                    // Insert the image from the memory stream.
                    Shape imageShape = builder.InsertImage(ms);
                    // Optionally set a size for visibility.
                    imageShape.Width = 100;
                    imageShape.Height = 100;
                }

                // Add a paragraph break after each image for spacing.
                builder.Writeln();
            }

            // Save the resulting document in the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
