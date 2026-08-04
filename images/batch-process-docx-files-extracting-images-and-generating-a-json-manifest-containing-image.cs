using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common
using Newtonsoft.Json;

public class Program
{
    // Manifest entry describing an extracted image.
    public class ImageInfo
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }

    public static void Main()
    {
        // Base directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "Output");
        string imagesDir = Path.Combine(outputDir, "ExtractedImages");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(imagesDir);

        // Create a deterministic sample image to be used in the documents.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a few sample DOCX files containing the sample image.
        CreateSampleDocuments(inputDir, sampleImagePath, 3);

        // List to hold manifest information.
        List<ImageInfo> manifest = new List<ImageInfo>();

        // Process each DOCX file in the input folder.
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to the output folder.
                shape.ImageData.Save(imageFullPath);

                // Record dimensions.
                ImageSize size = shape.ImageData.ImageSize;
                manifest.Add(new ImageInfo
                {
                    Document = Path.GetFileName(docPath),
                    ImageFile = imageFileName,
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels
                });

                imageIndex++;
            }

            // Validation: ensure at least one image was extracted from the document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images found in document '{docPath}'.");
        }

        // Validation: ensure the manifest contains entries.
        if (manifest.Count == 0)
            throw new InvalidOperationException("No images were extracted from any document.");

        // Serialize manifest to JSON.
        string manifestJson = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        string manifestPath = Path.Combine(outputDir, "manifest.json");
        File.WriteAllText(manifestPath, manifestJson);

        Console.WriteLine($"Processing complete. Extracted images saved to '{imagesDir}'.");
        Console.WriteLine($"Manifest written to '{manifestPath}'.");
    }

    // Creates a simple white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            bitmap.Save(filePath);
        }
    }

    // Generates a number of DOCX files, each containing the sample image.
    private static void CreateSampleDocuments(string folderPath, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            // Insert the sample image three times to have multiple images per file.
            for (int j = 0; j < 3; j++)
            {
                builder.InsertImage(imagePath);
                builder.Writeln(); // Add a line break between images.
            }

            string docFileName = $"SampleDocument{i}.docx";
            string docFullPath = Path.Combine(folderPath, docFileName);
            doc.Save(docFullPath);
        }
    }
}
