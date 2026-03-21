using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImagePreviewGenerator
{
    static void Main()
    {
        // Determine paths.
        string sourceDocPath = Path.Combine(Path.GetTempPath(), "input.docx");
        string previewFolder = Path.Combine(Path.GetTempPath(), "Previews");

        // Ensure the preview folder exists.
        Directory.CreateDirectory(previewFolder);

        Document doc;

        // If the source document does not exist, create a simple one with an embedded image.
        if (!File.Exists(sourceDocPath))
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A minimal 1x1 PNG image (transparent).
            byte[] pngData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            using (MemoryStream ms = new MemoryStream(pngData))
            {
                // Insert the image into the document.
                builder.InsertImage(ms);
            }

            // Save the temporary document.
            doc.Save(sourceDocPath);
        }
        else
        {
            doc = new Document(sourceDocPath);
        }

        int imageIndex = 0;

        // Iterate over all Shape nodes (including those in headers/footers).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    Scale = 0.75f // 75% of original size.
                };

                string outputPath = Path.Combine(previewFolder, $"image_{imageIndex}.png");
                shape.GetShapeRenderer().Save(outputPath, saveOptions);
                Console.WriteLine($"Saved preview: {outputPath}");
                imageIndex++;
            }
        }

        if (imageIndex == 0)
        {
            Console.WriteLine("No images were found in the document.");
        }
    }
}
