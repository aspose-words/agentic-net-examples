using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string imagesDir = Path.Combine(workDir, "Images");
        Directory.CreateDirectory(imagesDir);
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample image (PNG) using Aspose.Drawing
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // 2. Create a DOCX document and insert the sample image
        string docPath = Path.Combine(workDir, "SampleDocument.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the DOCX document
        Document doc = new Document(docPath);

        // 4. Extract images and collect metadata
        var metadataList = new List<ImageMeta>();
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedFileName = $"Image_{imageIndex}{extension}";
                string extractedPath = Path.Combine(outputDir, extractedFileName);

                // Save the image to disk
                shape.ImageData.Save(extractedPath);

                // Gather metadata
                var size = shape.ImageData.ImageSize;
                var meta = new ImageMeta
                {
                    Index = imageIndex,
                    FileName = extractedFileName,
                    ImageType = shape.ImageData.ImageType.ToString(),
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels,
                    HorizontalResolution = size.HorizontalResolution,
                    VerticalResolution = size.VerticalResolution
                };
                metadataList.Add(meta);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (metadataList.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 5. Generate a CSV file that can be opened by Excel
        string csvPath = Path.Combine(outputDir, "ImageMetadata.csv");
        WriteCsv(csvPath, metadataList);

        // Optional: also write metadata as JSON (demonstrates Newtonsoft.Json usage)
        string jsonPath = Path.Combine(outputDir, "ImageMetadata.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(metadataList, Formatting.Indented));

        // Indicate completion (no interactive prompts)
        Console.WriteLine($"Extraction complete. Images and metadata saved to: {outputDir}");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle
        var pen = new Pen(Color.Blue, 5);
        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        // Save and clean up
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX document and inserts the specified image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with sample image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Writes metadata to a CSV file
    private static void WriteCsv(string csvPath, List<ImageMeta> items)
    {
        using (var writer = new StreamWriter(csvPath, false))
        {
            // Header
            writer.WriteLine("Index,FileName,ImageType,WidthPixels,HeightPixels,HorizontalResolution,VerticalResolution");
            // Rows
            foreach (var item in items)
            {
                writer.WriteLine($"{item.Index},{item.FileName},{item.ImageType},{item.WidthPixels},{item.HeightPixels},{item.HorizontalResolution},{item.VerticalResolution}");
            }
        }
    }

    // Simple DTO for image metadata
    private class ImageMeta
    {
        public int Index { get; set; }
        public string FileName { get; set; }
        public string ImageType { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
        public double HorizontalResolution { get; set; }
        public double VerticalResolution { get; set; }
    }
}
