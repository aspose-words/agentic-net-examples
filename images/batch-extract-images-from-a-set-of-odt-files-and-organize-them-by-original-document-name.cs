using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define working folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputRootDir = Path.Combine(baseDir, "ExtractedImages");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputRootDir);

        // ------------------------------------------------------------
        // Step 1: Create a deterministic sample image (required by rules).
        // ------------------------------------------------------------
        CreateSampleImage(sampleImagePath);

        // ------------------------------------------------------------
        // Step 2: Generate a few ODT documents that contain the sample image.
        // ------------------------------------------------------------
        CreateSampleOdtDocument(Path.Combine(inputDir, "Doc1.odt"), sampleImagePath);
        CreateSampleOdtDocument(Path.Combine(inputDir, "Doc2.odt"), sampleImagePath);

        // ------------------------------------------------------------
        // Step 3: Batch process all ODT files, extracting images.
        // ------------------------------------------------------------
        string[] odtFiles = Directory.GetFiles(inputDir, "*.odt");
        if (odtFiles.Length == 0)
            throw new InvalidOperationException("No ODT files found for processing.");

        foreach (string odtFile in odtFiles)
        {
            // Load the document.
            Document doc = new Document(odtFile);

            // Collect all shape nodes that actually contain images.
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            if (shapeNodes.Count == 0)
                throw new InvalidOperationException($"Document '{Path.GetFileName(odtFile)}' contains no images.");

            // Prepare output folder for this document.
            string docNameWithoutExt = Path.GetFileNameWithoutExtension(odtFile);
            string docOutputDir = Path.Combine(outputRootDir, docNameWithoutExt);
            Directory.CreateDirectory(docOutputDir);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine file extension based on the image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(docOutputDir, $"image{imageIndex}{extension}");

                // Save the image to disk.
                shape.ImageData.Save(outFile);
                imageIndex++;
            }

            // Validation: ensure at least one image was saved.
            if (imageIndex == 0)
                throw new InvalidOperationException($"Failed to extract images from '{odtFile}'.");
        }

        // Program completed successfully.
    }

    // Creates a simple PNG image using Aspose.Drawing and saves it to the given path.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 100;

        // Create bitmap and draw deterministic content.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            graphics.Clear(Aspose.Drawing.Color.White);

            // Draw a blue rectangle.
            using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 3))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates an ODT document containing the specified image.
    private static void CreateSampleOdtDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph and the image.
        builder.Writeln("This document contains an image:");
        builder.InsertImage(imagePath);

        // Save as ODT.
        doc.Save(docPath, SaveFormat.Odt);
    }
}
