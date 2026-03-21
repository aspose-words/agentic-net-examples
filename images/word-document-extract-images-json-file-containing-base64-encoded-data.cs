using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageExtractionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Use relative paths so the example works out‑of‑the‑box.
            string inputFilePath = Path.Combine(AppContext.BaseDirectory, "Document.docx");
            string outputJsonPath = Path.Combine(AppContext.BaseDirectory, "ImagesBase64.json");

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath)!);

            // If the input file does not exist, create an empty document to avoid an exception.
            if (!File.Exists(inputFilePath))
            {
                var emptyDoc = new Document();
                emptyDoc.Save(inputFilePath);
            }

            // Load the Word document.
            Document doc = new Document(inputFilePath);

            // Collect all shapes in the document (including those inside headers/footers, tables, etc.).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            var images = new List<ImageInfo>();
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    byte[] imageBytes = shape.ImageData.ImageBytes;
                    string base64 = Convert.ToBase64String(imageBytes);

                    images.Add(new ImageInfo
                    {
                        Index = imageIndex,
                        Base64Data = base64,
                        ImageType = shape.ImageData.ImageType.ToString()
                    });

                    imageIndex++;
                }
            }

            // Serialize the list to JSON.
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(images, jsonOptions);

            // Write the JSON string to the output file.
            File.WriteAllText(outputJsonPath, json, Encoding.UTF8);

            Console.WriteLine($"Extraction complete. JSON written to: {outputJsonPath}");
        }

        private class ImageInfo
        {
            public int Index { get; set; }
            public string Base64Data { get; set; } = string.Empty;
            public string ImageType { get; set; } = string.Empty;
        }
    }
}
