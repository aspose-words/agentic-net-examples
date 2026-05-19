using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");

        // Clean previous runs.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images of different formats.
        CreateSampleImage(Path.Combine(baseDir, "sample1.png"), 200, 150, Color.LightBlue, ImageFormat.Png);
        CreateSampleImage(Path.Combine(baseDir, "sample2.jpg"), 180, 120, Color.LightCoral, ImageFormat.Jpeg);
        CreateSampleImage(Path.Combine(baseDir, "sample3.bmp"), 160, 100, Color.LightGreen, ImageFormat.Bmp);
        CreateSampleImage(Path.Combine(baseDir, "sample4.gif"), 140, 90, Color.LightYellow, ImageFormat.Gif);

        // Build a sample document containing the images.
        string docPath = Path.Combine(inputDir, "Sample.docx");
        BuildSampleDocument(docPath,
            Path.Combine(baseDir, "sample1.png"),
            Path.Combine(baseDir, "sample2.jpg"),
            Path.Combine(baseDir, "sample3.bmp"),
            Path.Combine(baseDir, "sample4.gif"));

        // Batch process all DOC/DOCX files in the input folder.
        int totalExtracted = 0;
        foreach (string file in Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly)
                                         .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                     f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
        {
            Document doc = new Document(file);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string formatFolder = extension.TrimStart('.').ToLowerInvariant(); // folder name based on format
                string targetFolder = Path.Combine(outputDir, formatFolder);
                Directory.CreateDirectory(targetFolder);

                string outputFileName = $"{Path.GetFileNameWithoutExtension(file)}_img{imageIndex}{extension}";
                string outputPath = Path.Combine(targetFolder, outputFileName);

                shape.ImageData.Save(outputPath);
                imageIndex++;
                totalExtracted++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        Console.WriteLine($"Extraction complete. Total images extracted: {totalExtracted}");
    }

    // Creates a deterministic bitmap and saves it using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Color backColor, ImageFormat format)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            // Simple visual cue: draw a diagonal line.
            graphics.DrawLine(new Pen(Color.Black, 2), 0, 0, width, height);
            bitmap.Save(filePath, format);
        }
    }

    // Builds a document and inserts the provided image files.
    private static void BuildSampleDocument(string docPath, params string[] imageFiles)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string img in imageFiles)
        {
            if (!File.Exists(img))
                throw new FileNotFoundException($"Image file not found: {img}");

            // Insert image inline.
            Shape shape = builder.InsertImage(img);
            shape.WrapType = WrapType.Inline;
            builder.Writeln(); // separate images with a line break.
        }

        doc.Save(docPath);
    }
}
