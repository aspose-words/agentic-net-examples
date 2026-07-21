using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

namespace AsposeWordsImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // -----------------------------------------------------------------
            // 1. Create a deterministic sample image (input.png).
            // -----------------------------------------------------------------
            string imagePath = Path.Combine(artifactsDir, "input.png");
            CreateSampleImage(imagePath);

            // -----------------------------------------------------------------
            // 2. Build a sample Word document that contains the image.
            // -----------------------------------------------------------------
            string docPath = Path.Combine(artifactsDir, "sample.docx");
            CreateSampleDocument(docPath, imagePath);

            // -----------------------------------------------------------------
            // 3. Load the document and extract all images.
            // -----------------------------------------------------------------
            Document doc = new Document(docPath);
            List<ImageInfo> extractedImages = ExtractImages(doc);

            // Validate that at least one image was extracted.
            if (extractedImages.Count == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // -----------------------------------------------------------------
            // 4. Serialize the extracted images to JSON (base64 encoded).
            // -----------------------------------------------------------------
            string json = JsonConvert.SerializeObject(extractedImages, Formatting.Indented);
            string jsonPath = Path.Combine(artifactsDir, "images.json");
            File.WriteAllText(jsonPath, json);

            // Verify that the JSON file was created.
            if (!File.Exists(jsonPath))
                throw new InvalidOperationException("Failed to create the JSON output file.");

            // The example finishes here; no interactive prompts are used.
        }

        private static void CreateSampleImage(string filePath)
        {
            // Create a 100x100 white bitmap using Aspose.Drawing.
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                }

                // Save the bitmap to a deterministic file name.
                bitmap.Save(filePath);
            }
        }

        private static void CreateSampleDocument(string docPath, string imagePath)
        {
            // Create a blank document and insert the sample image twice.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Sample document with images:");
            builder.InsertImage(imagePath);
            builder.Writeln();
            builder.InsertImage(imagePath);

            // Save the document.
            doc.Save(docPath);
        }

        private static List<ImageInfo> ExtractImages(Document doc)
        {
            var images = new List<ImageInfo>();
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int index = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Save the image data to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset before reading.

                    byte[] imageBytes = ms.ToArray();
                    string base64 = Convert.ToBase64String(imageBytes);
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    images.Add(new ImageInfo
                    {
                        Index = index,
                        Extension = extension,
                        Base64 = base64
                    });

                    index++;
                }
            }

            return images;
        }

        // Helper class for JSON serialization.
        private class ImageInfo
        {
            public int Index { get; set; }
            public string Extension { get; set; }
            public string Base64 { get; set; }
        }
    }
}
