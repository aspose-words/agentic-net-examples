using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OdtImageExtractor
{
    // Simple DTO for JSON manifest entries
    public class ImageManifestEntry
    {
        public string SourceDocument { get; set; }   // ODT file name
        public string ImageFile { get; set; }        // Extracted image file name (relative to manifest)
    }

    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Folder containing ODT files to process
            string odtFolder = @"C:\InputOdtFiles";

            // Folder where extracted images will be saved
            string imagesOutputFolder = @"C:\ExtractedImages";

            // Path for the consolidated JSON manifest
            string manifestPath = @"C:\ImageManifest.json";

            // Ensure the images output folder exists
            if (!Directory.Exists(imagesOutputFolder))
                Directory.CreateDirectory(imagesOutputFolder);

            // List that will hold all manifest entries
            List<ImageManifestEntry> manifest = new List<ImageManifestEntry>();

            // Process each ODT file in the input folder
            foreach (string odtFilePath in Directory.GetFiles(odtFolder, "*.odt"))
            {
                // Load the ODT document using the provided Document constructor
                Document doc = new Document(odtFilePath);

                // Retrieve all Shape nodes (including pictures) from the document
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes)
                {
                    // Only process shapes that actually contain an image
                    if (!shape.HasImage)
                        continue;

                    // Build a unique image file name: <odtFileName>_img<index>.<extension>
                    string odtFileName = Path.GetFileNameWithoutExtension(odtFilePath);
                    string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{odtFileName}_img{imageIndex}{imageExtension}";
                    string imageFullPath = Path.Combine(imagesOutputFolder, imageFileName);

                    // Save the image using the provided ImageData.Save method
                    shape.ImageData.Save(imageFullPath);

                    // Add an entry to the manifest
                    manifest.Add(new ImageManifestEntry
                    {
                        SourceDocument = Path.GetFileName(odtFilePath),
                        ImageFile = imageFileName
                    });

                    imageIndex++;
                }
            }

            // Serialize the manifest list to JSON using System.Text.Json (standard .NET library)
            string json = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON manifest to disk
            File.WriteAllText(manifestPath, json);
        }
    }
}
