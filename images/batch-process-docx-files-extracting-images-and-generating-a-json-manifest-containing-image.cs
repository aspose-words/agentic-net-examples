using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string imagesDir = Path.Combine(baseDir, "Images");
        string docsDir = Path.Combine(baseDir, "Docs");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        if (!File.Exists(sampleImagePath))
        {
            // Create a 200x200 white bitmap
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                }
                bitmap.Save(sampleImagePath);
            }
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image
        // -------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            string docPath = Path.Combine(docsDir, $"Doc{docIndex}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the sample image twice into each document
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // add a line break
            builder.InsertImage(sampleImagePath);

            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process DOCX files: extract images and build manifest
        // -------------------------------------------------
        List<DocumentManifest> manifest = new List<DocumentManifest>();

        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            List<ImageInfo> imagesInfo = new List<ImageInfo>();
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);

                // Save the image to disk
                shape.ImageData.Save(imageFullPath);

                // Retrieve image dimensions
                ImageSize size = shape.ImageData.ImageSize;
                imagesInfo.Add(new ImageInfo
                {
                    FileName = imageFileName,
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels
                });

                imageIndex++;
            }

            // Validation: each document must contain at least one extracted image
            if (imagesInfo.Count == 0)
                throw new InvalidOperationException($"No images were extracted from document '{Path.GetFileName(docFile)}'.");

            manifest.Add(new DocumentManifest
            {
                DocumentName = Path.GetFileName(docFile),
                Images = imagesInfo
            });
        }

        // -------------------------------------------------
        // 4. Serialize manifest to JSON
        // -------------------------------------------------
        string manifestJson = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
        string manifestPath = Path.Combine(outputDir, "manifest.json");
        File.WriteAllText(manifestPath, manifestJson);
    }

    // Helper classes for JSON manifest
    public class ImageInfo
    {
        public string FileName { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }

    public class DocumentManifest
    {
        public string DocumentName { get; set; }
        public List<ImageInfo> Images { get; set; }
    }
}
