using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtraction");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a deterministic sample image (sample.png) using Aspose.Drawing
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // -----------------------------------------------------------------
        // Create a sample DOCX file that contains the image
        // -----------------------------------------------------------------
        string sampleDocPath = Path.Combine(inputDir, "SampleDocument.docx");
        CreateSampleDocumentWithImage(sampleDocPath, sampleImagePath);

        // -----------------------------------------------------------------
        // Batch process all DOC/DOCX files in the input folder
        // -----------------------------------------------------------------
        var docFiles = Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly)
                                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));

        foreach (var docFile in docFiles)
        {
            // Load the document
            Document doc = new Document(docFile, new LoadOptions());

            // Get all shape nodes that contain images
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            if (!shapes.Any())
                throw new InvalidOperationException($"No images found in document '{Path.GetFileName(docFile)}'.");

            int imageIndex = 0;
            foreach (var shape in shapes)
            {
                // Retrieve raw image bytes
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load the bytes into an Aspose.Drawing.Bitmap
                using (var ms = new MemoryStream(imageBytes))
                using (var bitmap = new Bitmap(ms))
                {
                    // Prepare output file name (always BMP)
                    string outputFileName = $"Img_{Path.GetFileNameWithoutExtension(docFile)}_{imageIndex}.bmp";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    // Save as BMP
                    bitmap.Save(outputPath, ImageFormat.Bmp);
                }

                imageIndex++;
            }
        }

        // Validate that at least one BMP file was created
        var bmpFiles = Directory.GetFiles(outputDir, "*.bmp");
        if (!bmpFiles.Any())
            throw new InvalidOperationException("No BMP images were extracted.");

        Console.WriteLine("Image extraction completed successfully.");
    }

    // Creates a simple PNG image with a white background and a black rectangle.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (var bitmap = new Bitmap(width, height))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple black rectangle
            var pen = new Pen(Color.Black, 3);
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            pen.Dispose();

            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX document and inserts the specified image.
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);
    }
}
