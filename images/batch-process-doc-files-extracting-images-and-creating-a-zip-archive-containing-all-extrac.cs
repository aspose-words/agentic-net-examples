using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

// Alias to disambiguate the CompressionLevel enum used for ZIP creation.
using SysCompressionLevel = System.IO.Compression.CompressionLevel;

public class Program
{
    public static void Main()
    {
        // Root folder for all temporary data
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "ImageBatchDemo");
        string docsFolder = Path.Combine(rootFolder, "Docs");
        string imagesFolder = Path.Combine(rootFolder, "ExtractedImages");
        string zipFolder = Path.Combine(rootFolder, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(docsFolder);
        Directory.CreateDirectory(imagesFolder);
        Directory.CreateDirectory(zipFolder);

        // -------------------------------------------------
        // 1. Create deterministic sample images (PNG)
        // -------------------------------------------------
        string sampleImagePath1 = Path.Combine(rootFolder, "sample1.png");
        string sampleImagePath2 = Path.Combine(rootFolder, "sample2.png");
        CreateSamplePng(sampleImagePath1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng(sampleImagePath2, 150, 200, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images
        // -------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert one of the sample images into the document
            string imagePath = i == 1 ? sampleImagePath1 : sampleImagePath2;
            builder.InsertImage(imagePath);

            // Save the document
            string docPath = Path.Combine(docsFolder, $"SampleDoc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process each DOCX, extract images to folder
        // -------------------------------------------------
        int totalExtracted = 0;
        string[] docFiles = Directory.GetFiles(docsFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string extractedFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                    string extractedPath = Path.Combine(imagesFolder, extractedFileName);

                    // Save the image
                    shape.ImageData.Save(extractedPath);
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 4. Create a ZIP archive containing all extracted images
        // -------------------------------------------------
        string zipPath = Path.Combine(zipFolder, "ExtractedImages.zip");
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // Use the aliased CompressionLevel to avoid ambiguity
        ZipFile.CreateFromDirectory(imagesFolder, zipPath, SysCompressionLevel.Optimal, includeBaseDirectory: false);

        // Simple verification output
        Console.WriteLine($"Processed {docFiles.Length} document(s).");
        Console.WriteLine($"Extracted {totalExtracted} image(s) to \"{imagesFolder}\".");
        Console.WriteLine($"Created ZIP archive at \"{zipPath}\".");
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing
    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            // Optionally draw a simple rectangle border
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
            {
                graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
            }
            // Save the bitmap as PNG
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
