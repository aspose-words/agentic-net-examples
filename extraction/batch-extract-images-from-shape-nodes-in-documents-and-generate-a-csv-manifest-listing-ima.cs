using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

namespace ImageBatchExtraction
{
    public class Program
    {
        // Base64-encoded 1x1 pixel PNG used for sample images.
        private const string SamplePngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";

        public static void Main()
        {
            // Prepare folders.
            string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Create sample documents with images.
            CreateSampleDocument(Path.Combine(inputFolder, "DocA.docx"), "ImageA");
            CreateSampleDocument(Path.Combine(inputFolder, "DocB.docx"), "ImageB");

            // List to hold manifest rows.
            List<string> manifestLines = new List<string>
            {
                "ImageFile,SourceDocument"
            };

            // Process each document in the input folder.
            foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                Document loaded = new Document(docPath);
                // Find all shape nodes that contain images.
                var imageShapes = loaded.GetChildNodes(NodeType.Shape, true)
                                        .OfType<Shape>()
                                        .Where(s => s.HasImage)
                                        .ToList();

                if (imageShapes.Count == 0)
                    continue; // No images in this document.

                int imageIndex = 0;
                foreach (Shape shape in imageShapes)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}.png";
                    string imageFilePath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imageFilePath);
                    manifestLines.Add($"{imageFileName},{Path.GetFileName(docPath)}");
                    imageIndex++;
                }
            }

            // Validate that at least one image was extracted.
            if (manifestLines.Count <= 1)
                throw new InvalidOperationException("No images were extracted from the batch.");

            // Write CSV manifest.
            string manifestPath = Path.Combine(outputFolder, "manifest.csv");
            File.WriteAllLines(manifestPath, manifestLines);

            // Optional: Verify manifest file creation.
            if (!File.Exists(manifestPath))
                throw new InvalidOperationException("Manifest CSV was not created.");
        }

        // Creates a simple DOCX file containing a single image shape.
        private static void CreateSampleDocument(string filePath, string bookmarkName)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark to identify the document (not required for extraction but adds variety).
            builder.StartBookmark(bookmarkName);
            builder.Writeln($"This is {bookmarkName} document.");
            builder.EndBookmark(bookmarkName);

            // Insert the sample image.
            byte[] pngBytes = Convert.FromBase64String(SamplePngBase64);
            using (MemoryStream imageStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imageStream);
            }

            doc.Save(filePath);
        }
    }
}
