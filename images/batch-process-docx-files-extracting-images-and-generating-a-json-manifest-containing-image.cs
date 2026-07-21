using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders for input documents, extracted images and the JSON manifest.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "ImageManifest.json");

        // Ensure clean environment.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        CleanDirectory(inputDir);
        CleanDirectory(imagesDir);
        if (File.Exists(manifestPath)) File.Delete(manifestPath);

        // -------------------------------------------------
        // 1. Create deterministic sample images.
        // -------------------------------------------------
        string pngPath = Path.Combine(baseDir, "sample1.png");
        string jpgPath = Path.Combine(baseDir, "sample2.jpg");

        CreateSampleImage(pngPath, 200, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(jpgPath, 300, 150, Aspose.Drawing.Color.LightCoral, Aspose.Drawing.Imaging.ImageFormat.Jpeg);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {docIndex}");
            builder.InsertImage(pngPath);
            builder.InsertImage(jpgPath);
            string docPath = Path.Combine(inputDir, $"Doc{docIndex}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images and build manifest.
        // -------------------------------------------------
        var manifest = new List<ImageManifestEntry>();
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx", SearchOption.TopDirectoryOnly);

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine file extension based on the image type stored in the shape.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save extracted image: {imageFullPath}");

                // Retrieve image dimensions.
                ImageSize size = shape.ImageData.ImageSize;

                // Add entry to the manifest.
                manifest.Add(new ImageManifestEntry
                {
                    Document = Path.GetFileName(docFile),
                    ImageFile = imageFileName,
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels
                });

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (manifest.Count == 0)
            throw new InvalidOperationException("No images were extracted from the input documents.");

        // -------------------------------------------------
        // 4. Serialize manifest to JSON.
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);
    }

    // Helper to create a deterministic bitmap and save it to a file.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor, Aspose.Drawing.Imaging.ImageFormat format = null)
    {
        // Default to PNG if no format is supplied.
        if (format == null) format = Aspose.Drawing.Imaging.ImageFormat.Png;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            // Optionally draw a simple rectangle to make the image non‑empty.
            graphics.DrawRectangle(new Pen(Aspose.Drawing.Color.Black), 0, 0, width - 1, height - 1);
            bitmap.Save(filePath, format);
        }

        // Verify that the file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }

    // Helper to delete all files in a directory (keeps the directory itself).
    private static void CleanDirectory(string directoryPath)
    {
        if (!Directory.Exists(directoryPath)) return;
        foreach (string file in Directory.GetFiles(directoryPath))
        {
            File.Delete(file);
        }
    }

    // Manifest entry model.
    private class ImageManifestEntry
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }
}
