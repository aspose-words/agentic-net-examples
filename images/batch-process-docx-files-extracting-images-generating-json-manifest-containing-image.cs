using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace BatchImageExtractor
{
    // Simple DTO for JSON manifest entries
    public class ImageManifestEntry
    {
        public string DocumentPath { get; set; }
        public string ImageFileName { get; set; }
        public double WidthPoints { get; set; }
        public double HeightPoints { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing DOCX files
            string inputFolder = @"C:\Docs\Input";
            // Folder where extracted images will be saved
            string imagesOutputFolder = @"C:\Docs\ExtractedImages";
            // Path for the JSON manifest file
            string manifestPath = @"C:\Docs\image-manifest.json";

            // Ensure output directories exist
            Directory.CreateDirectory(imagesOutputFolder);

            // List to hold manifest entries
            List<ImageManifestEntry> manifest = new List<ImageManifestEntry>();

            // Process each .docx file in the input folder
            foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load the document using the provided constructor
                Document doc = new Document(docPath);

                // Get all Shape nodes (including images) in the document
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes)
                {
                    if (!shape.HasImage)
                        continue; // Skip non‑image shapes

                    // Build a unique image file name
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    string imageFullPath = Path.Combine(imagesOutputFolder, imageFileName);

                    // Save the image bytes to disk
                    shape.ImageData.Save(imageFullPath);

                    // Record dimensions (Width and Height are in points)
                    manifest.Add(new ImageManifestEntry
                    {
                        DocumentPath = docPath,
                        ImageFileName = imageFullPath,
                        WidthPoints = shape.Width,
                        HeightPoints = shape.Height
                    });

                    imageIndex++;
                }
            }

            // Serialize manifest to JSON with indentation for readability
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(manifest, jsonOptions);
            File.WriteAllText(manifestPath, json);
        }
    }
}
