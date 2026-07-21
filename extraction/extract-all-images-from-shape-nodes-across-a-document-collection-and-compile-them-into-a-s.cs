using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "Images.zip");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);

        // Sample PNG (1x1 pixel) in Base64.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";

        // Create a few sample documents each containing an image.
        for (int docIndex = 0; docIndex < 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (MemoryStream imgStream = new MemoryStream(Convert.FromBase64String(pngBase64)))
            {
                builder.InsertImage(imgStream);
            }

            string docPath = Path.Combine(inputDir, $"Sample{docIndex}.docx");
            doc.Save(docPath);
        }

        // Extract images from all documents in the collection.
        int globalImageIndex = 0;
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document loadedDoc = new Document(docPath);
            var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Where(s => s.HasImage);

            foreach (Shape shape in shapes)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"img_{globalImageIndex}{extension}";
                string imageFullPath = Path.Combine(imageDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                globalImageIndex++;
            }
        }

        // Verify that at least one image was extracted.
        if (!Directory.GetFiles(imageDir).Any())
            throw new InvalidOperationException("No images were extracted.");

        // Create a ZIP archive containing all extracted images.
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        ZipFile.CreateFromDirectory(imageDir, zipPath);

        // Validate ZIP creation.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        // Optional: clean up temporary folders (comment out if inspection is needed).
        // Directory.Delete(inputDir, true);
        // Directory.Delete(imageDir, true);
    }
}
