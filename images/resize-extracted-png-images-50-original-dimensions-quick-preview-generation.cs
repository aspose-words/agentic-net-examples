using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ResizeExtractedPngImages
{
    static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "ResizeExtractedPngImages");
        Directory.CreateDirectory(workFolder);

        // Create a tiny PNG file (1×1 pixel, red) if it does not already exist.
        string pngPath = Path.Combine(workFolder, "sample.png");
        if (!File.Exists(pngPath))
        {
            // Base64‑encoded PNG (red 1×1 pixel).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK7cAAAAASUVORK5CYII=";
            File.WriteAllBytes(pngPath, Convert.FromBase64String(base64Png));
        }

        // Create a Word document and insert the PNG image.
        string docPath = Path.Combine(workFolder, "DocumentWithImages.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        doc.Save(docPath);

        // Load the document (simulating an external source).
        Document loadedDoc = new Document(docPath);

        // Prepare a folder for the resized preview images.
        string previewFolder = Path.Combine(workFolder, "Previews");
        Directory.CreateDirectory(previewFolder);

        // Iterate over all Shape nodes in the document.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage) continue;

            // Ensure the image is a PNG; skip other formats.
            string imageExtension = shape.ImageData.ImageType.ToString().ToLower();
            if (imageExtension != "png") continue;

            // Build a file name for the preview image.
            string previewPath = Path.Combine(previewFolder, $"ImagePreview_{imageIndex++}.png");

            // Configure ImageSaveOptions to render the shape at 50% scale.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Scale = 0.5f
            };

            // Render the shape to a PNG file using the configured options.
            shape.GetShapeRenderer().Save(previewPath, saveOptions);
        }

        Console.WriteLine($"Preview images have been generated in: {previewFolder}");
    }
}
