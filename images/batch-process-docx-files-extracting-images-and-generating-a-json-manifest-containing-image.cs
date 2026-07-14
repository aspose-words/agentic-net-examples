using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

namespace BatchImageExtraction
{
    // Represents a single image entry in the JSON manifest.
    public class ImageManifestEntry
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Base working directory.
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImages");
            Directory.CreateDirectory(baseDir);

            // Folder for extracted images.
            string imagesDir = Path.Combine(baseDir, "ExtractedImages");
            Directory.CreateDirectory(imagesDir);

            // Create a deterministic sample image (sample.png).
            string sampleImagePath = Path.Combine(baseDir, "sample.png");
            CreateSampleImage(sampleImagePath, 100, 100);

            // Create sample DOCX files containing the sample image.
            CreateSampleDocument(Path.Combine(baseDir, "Doc1.docx"), sampleImagePath, 2);
            CreateSampleDocument(Path.Combine(baseDir, "Doc2.docx"), sampleImagePath, 3);

            // List to hold manifest entries.
            List<ImageManifestEntry> manifest = new List<ImageManifestEntry>();

            // Process each DOCX file in the batch folder.
            foreach (string docPath in Directory.GetFiles(baseDir, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(docPath);

                // Retrieve all shape nodes.
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;

                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage) continue;

                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image_{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imageFullPath);

                    // Capture image dimensions.
                    ImageSize size = shape.ImageData.ImageSize;
                    manifest.Add(new ImageManifestEntry
                    {
                        Document = Path.GetFileName(docPath),
                        ImageFile = imageFileName,
                        WidthPixels = size.WidthPixels,
                        HeightPixels = size.HeightPixels
                    });

                    imageIndex++;
                }
            }

            // Validate that at least one image was extracted.
            if (manifest.Count == 0)
                throw new InvalidOperationException("No images were extracted from the batch documents.");

            // Serialize manifest to JSON.
            string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
            string manifestPath = Path.Combine(baseDir, "manifest.json");
            File.WriteAllText(manifestPath, json);
        }

        // Creates a deterministic PNG image using Aspose.Drawing.
        private static void CreateSampleImage(string filePath, int width, int height)
        {
            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                bitmap.Save(filePath, ImageFormat.Png);
            }
        }

        // Creates a DOCX file with a specified number of inserted images.
        private static void CreateSampleDocument(string docPath, string imagePath, int imageCount)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < imageCount; i++)
            {
                builder.InsertParagraph();
                builder.InsertImage(imagePath);
            }

            doc.Save(docPath, SaveFormat.Docx);
        }
    }
}
