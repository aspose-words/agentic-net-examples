using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Root folder for the example
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImages");
        string inputFolder = Path.Combine(rootFolder, "Input");
        string outputFolder = Path.Combine(rootFolder, "ExtractedImages");
        string manifestPath = Path.Combine(rootFolder, "manifest.json");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (200x200 white PNG)
        string sampleImagePath = Path.Combine(rootFolder, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a few DOCX files that contain the sample image
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Document{i}.docx");
            CreateDocumentWithImages(docPath, sampleImagePath, i);
        }

        // List to hold manifest entries
        var manifest = new List<ImageInfo>();

        // Process each DOCX file in the input folder
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputFolder, imageFileName);

                // Save the image to disk
                shape.ImageData.Save(imageFullPath);

                // Retrieve image dimensions
                ImageSize size = shape.ImageData.ImageSize;
                manifest.Add(new ImageInfo
                {
                    Document = Path.GetFileName(docFile),
                    ImageFile = imageFileName,
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels
                });

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (!manifest.Any())
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Serialize manifest to JSON
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);
    }

    // Creates a simple white PNG image using Aspose.Drawing
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a DOCX file and inserts the sample image a number of times
    private static void CreateDocumentWithImages(string docPath, string imagePath, int repeatCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < repeatCount; i++)
        {
            builder.Writeln($"Image insertion #{i + 1}");
            builder.InsertImage(imagePath);
            builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(docPath);
    }

    // DTO for JSON manifest entries
    private class ImageInfo
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }
}
