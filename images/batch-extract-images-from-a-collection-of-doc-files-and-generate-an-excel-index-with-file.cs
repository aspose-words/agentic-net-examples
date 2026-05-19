using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define working folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string indexCsvPath = Path.Combine(baseDir, "ImageIndex.csv");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (PNG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImage1Path = Path.Combine(baseDir, "sample1.png");
        string sampleImage2Path = Path.Combine(baseDir, "sample2.png");

        CreateSamplePng(sampleImage1Path, 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng(sampleImage2Path, 120, 80, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Create a few sample DOCX files that contain the images.
        // -----------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {docIndex}");
            // Insert first image.
            builder.InsertImage(sampleImage1Path);
            builder.Writeln(); // separate
            // Insert second image.
            builder.InsertImage(sampleImage2Path);

            string docPath = Path.Combine(inputDocsDir, $"SampleDocument{docIndex}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Batch extract images from all DOC/DOCX files.
        // -----------------------------------------------------------------
        var indexLines = new List<string>();
        int totalExtracted = 0;

        string[] docFiles = Directory.GetFiles(inputDocsDir, "*.*", SearchOption.TopDirectoryOnly)
                                     .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                 f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                                     .ToArray();

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine proper file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imageFullPath);

                    // Record entry for the index.
                    indexLines.Add($"{Path.GetFullPath(docFile)},{Path.GetFullPath(imageFullPath)}");
                    imageIndex++;
                    totalExtracted++;
                }
            }

            // Validation: ensure at least one image was found in the current document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{docFile}'.");
        }

        // -----------------------------------------------------------------
        // 4. Write the CSV index file (Excel can open CSV directly).
        // -----------------------------------------------------------------
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from any document.");

        // CSV header.
        indexLines.Insert(0, "DocumentPath,ImagePath");
        File.WriteAllLines(indexCsvPath, indexLines);

        // -----------------------------------------------------------------
        // 5. Simple validation that output files exist.
        // -----------------------------------------------------------------
        if (!File.Exists(indexCsvPath))
            throw new FileNotFoundException("Index CSV file was not created.", indexCsvPath);

        // All done – the program exits without waiting for user input.
    }

    // Helper method to create a PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
