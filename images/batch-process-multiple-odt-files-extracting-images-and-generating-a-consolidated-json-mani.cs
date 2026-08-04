using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json;

namespace BatchOdtImageExtractor
{
    public class Program
    {
        // DTO for JSON manifest
        public class DocumentManifest
        {
            public string DocumentName { get; set; }
            public List<string> ImageFiles { get; set; } = new List<string>();
        }

        static void Main()
        {
            // Define folders relative to the executable location
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputFolder = Path.Combine(baseDir, "InputDocs");
            string imageFolder = Path.Combine(baseDir, "ExtractedImages");
            string outputFolder = Path.Combine(baseDir, "Output");
            string sampleImagePath = Path.Combine(baseDir, "sample.png");

            // Ensure required directories exist
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(imageFolder);
            Directory.CreateDirectory(outputFolder);

            // Create a deterministic sample image (100x100 white background)
            CreateSampleImage(sampleImagePath, 100, 100);

            // Generate a few ODT documents each containing two inserted images
            CreateSampleOdtDocuments(inputFolder, sampleImagePath, 3);

            // Process each ODT file, extract images, and build manifest
            List<DocumentManifest> manifest = new List<DocumentManifest>();
            int totalExtractedImages = 0;

            foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
            {
                Document doc = new Document(odtPath);
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
                var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

                DocumentManifest docEntry = new DocumentManifest
                {
                    DocumentName = Path.GetFileName(odtPath)
                };

                int imageIndex = 0;
                foreach (Shape shape in imageShapes)
                {
                    // Determine file extension based on image type
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtPath)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imageFolder, imageFileName);

                    // Save the image to disk
                    shape.ImageData.Save(imageFullPath);
                    docEntry.ImageFiles.Add(imageFileName);
                    imageIndex++;
                    totalExtractedImages++;
                }

                manifest.Add(docEntry);
            }

            // Validate that at least one image was extracted
            if (totalExtractedImages == 0)
                throw new InvalidOperationException("No images were extracted from the ODT files.");

            // Serialize manifest to JSON
            string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
            string manifestPath = Path.Combine(outputFolder, "manifest.json");
            File.WriteAllText(manifestPath, json);
        }

        // Creates a simple PNG image using Aspose.Drawing
        private static void CreateSampleImage(string filePath, int width, int height)
        {
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
            {
                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Aspose.Drawing.Color.White);
                }
                bitmap.Save(filePath);
            }
        }

        // Generates a number of ODT documents each containing two inserted images
        private static void CreateSampleOdtDocuments(string folderPath, string imagePath, int count)
        {
            for (int i = 0; i < count; i++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                builder.Writeln($"Document {i + 1}");
                // Insert first image
                builder.InsertImage(imagePath);
                builder.Writeln(); // line break
                // Insert second image
                builder.InsertImage(imagePath);

                string odtFileName = Path.Combine(folderPath, $"SampleDocument_{i + 1}.odt");
                doc.Save(odtFileName, SaveFormat.Odt);
            }
        }
    }
}
