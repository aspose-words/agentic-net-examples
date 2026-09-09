using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class BatchImageExtractor
{
    // Entry for the JSON manifest.
    private class ManifestEntry
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }

    public static void Main()
    {
        // Base directories for input documents, extracted images and the manifest.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png) to be used.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Generate a few sample DOCX files that contain the image.
        // -----------------------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(baseDir, $"Doc{i}.docx");
            CreateSampleDocument(docPath, sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 3. Batch process all DOCX files: extract images and build manifest.
        // -----------------------------------------------------------------
        var manifest = new List<ManifestEntry>();
        string[] docFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.TopDirectoryOnly);
        foreach (string docFile in docFiles)
        {
            // Load the document.
            Document doc = new Document(docFile);

            // Collect all shape nodes that contain images.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);
                imageIndex++;

                // Retrieve image dimensions.
                ImageSize size = shape.ImageData.ImageSize;
                manifest.Add(new ManifestEntry
                {
                    Document = Path.GetFileName(docFile),
                    ImageFile = imageFileName,
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels
                });
            }

            // Validation: each document must contain at least one extracted image.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from '{docFile}'.");
        }

        // -----------------------------------------------------------------
        // 4. Serialize the manifest to JSON.
        // -----------------------------------------------------------------
        string manifestPath = Path.Combine(baseDir, "manifest.json");
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);
    }

    // Creates a simple white PNG image of the specified size.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file containing a single paragraph and the provided image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample document generated for '{Path.GetFileName(docPath)}'.");
        // Insert the image; the builder returns a Shape that already has the image.
        Shape imgShape = builder.InsertImage(imagePath);
        // Ensure the shape indeed has an image before proceeding.
        if (!imgShape.HasImage)
            throw new InvalidOperationException("Failed to insert image into the document.");
        doc.Save(docPath);
    }
}
