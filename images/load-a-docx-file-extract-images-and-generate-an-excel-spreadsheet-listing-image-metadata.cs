using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Build a DOCX document and insert the sample image twice
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateDocumentWithImages(docPath, sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Load the DOCX, extract all images and collect metadata
        // -----------------------------------------------------------------
        var imageInfos = new List<ImageInfo>();
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"Image_{imageIndex}{extension}";
            string imageFilePath = Path.Combine(outputDir, imageFileName);

            // Save the image to disk
            shape.ImageData.Save(imageFilePath);

            // Gather metadata
            var size = shape.ImageData.ImageSize;
            imageInfos.Add(new ImageInfo
            {
                Index = imageIndex,
                FileName = imageFileName,
                ImageType = shape.ImageData.ImageType.ToString(),
                WidthPixels = size.WidthPixels,
                HeightPixels = size.HeightPixels,
                HorizontalResolution = size.HorizontalResolution,
                VerticalResolution = size.VerticalResolution
            });

            imageIndex++;
        }

        // Validate that at least one image was extracted
        if (imageInfos.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Generate a CSV file that can be opened by Excel
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(outputDir, "ImageMetadata.csv");
        WriteCsv(csvPath, imageInfos);

        // -----------------------------------------------------------------
        // 5. Informative output (no interactive prompts)
        // -----------------------------------------------------------------
        Console.WriteLine($"Document created: {docPath}");
        Console.WriteLine($"Extracted {imageInfos.Count} image(s) to folder: {outputDir}");
        Console.WriteLine($"Metadata CSV generated: {csvPath}");
    }

    // Creates a simple white bitmap and saves it to the specified path
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Additional deterministic drawing can be added here if desired
            bitmap.Save(filePath);
        }
    }

    // Builds a DOCX file and inserts the provided image file twice
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with images:");
        builder.InsertImage(imagePath);
        builder.Writeln(); // add a line break
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }

    // Writes image metadata to a CSV file
    private static void WriteCsv(string csvPath, List<ImageInfo> infos)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Index,FileName,ImageType,WidthPixels,HeightPixels,HorizontalResolution,VerticalResolution");

        foreach (var info in infos)
        {
            sb.AppendLine($"{info.Index},{info.FileName},{info.ImageType},{info.WidthPixels},{info.HeightPixels},{info.HorizontalResolution},{info.VerticalResolution}");
        }

        File.WriteAllText(csvPath, sb.ToString());
    }

    // Simple DTO for image metadata
    private class ImageInfo
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
