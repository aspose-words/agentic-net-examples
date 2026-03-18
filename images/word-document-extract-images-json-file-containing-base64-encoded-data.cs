using System;
using System.Collections.Generic;
using System.IO;
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
            // Path to the source Word document.
            string inputFilePath = @"C:\Input\Document.docx";

            // Path where the JSON output will be written.
            string outputJsonPath = @"C:\Output\ImagesBase64.json";

            // Load the Word document using the constructor that accepts a file name.
            Document doc = new Document(inputFilePath);

            // Collect all shapes in the document (including those inside headers/footers, tables, etc.).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            // Prepare a list to hold image data objects.
            var images = new List<ImageInfo>();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                // Process only shapes that actually contain an image.
                if (shape.HasImage)
                {
                    // Retrieve the raw image bytes.
                    byte[] imageBytes = shape.ImageData.ImageBytes;

                    // Convert the bytes to a Base64 string.
                    string base64 = Convert.ToBase64String(imageBytes);

                    // Store the information in the list.
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
        }

        // Helper class representing a single image entry in the JSON output.
        private class ImageInfo
        {
            public int Index { get; set; }
            public string Base64Data { get; set; }
            public string ImageType { get; set; }
        }
    }
}
