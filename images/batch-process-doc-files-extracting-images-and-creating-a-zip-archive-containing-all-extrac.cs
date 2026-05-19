using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Base directories for the example.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string inputDocsDir = Path.Combine(artifactsDir, "InputDocs");
        string extractedImagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        string zipPath = Path.Combine(artifactsDir, "ExtractedImages.zip");

        // Ensure clean environment.
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir, true);
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create sample images using Aspose.Drawing.
        // -------------------------------------------------
        string[] sampleImagePaths = new string[2];
        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            string imagePath = Path.Combine(artifactsDir, $"sample{i + 1}.png");
            CreateSampleImage(imagePath, 200 + i * 50, 150 + i * 30);
            sampleImagePaths[i] = imagePath;
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -------------------------------------------------
        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i + 1} with an image:");
            builder.InsertImage(sampleImagePaths[i]);
            string docPath = Path.Combine(inputDocsDir, $"Document{i + 1}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOC/DOCX files, extract images.
        // -------------------------------------------------
        int totalExtracted = 0;
        foreach (string docFile in Directory.GetFiles(inputDocsDir, "*.doc*"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine a deterministic file name.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 4. Create a ZIP archive containing all extracted images.
        // -------------------------------------------------
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(extractedImagesDir, zipPath);

        // Simple verification that the ZIP file exists.
        if (!File.Exists(zipPath))
            throw new FileNotFoundException("Failed to create the ZIP archive.");

        // The example finishes here; no interactive prompts are used.
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap and draw a simple rectangle with a solid background.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a filled rectangle.
            graphics.FillRectangle(new SolidBrush(Color.FromArgb(100, 150, 200)), 0, 0, width, height);
            // Save the bitmap as PNG.
            bitmap.Save(filePath);
        }
    }
}
