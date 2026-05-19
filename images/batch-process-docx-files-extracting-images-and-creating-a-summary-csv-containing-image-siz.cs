using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Bitmap, Graphics, Color
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string sampleImageDir = Path.Combine(baseDir, "SampleImages");
        string extractedImageDir = Path.Combine(baseDir, "ExtractedImages");
        string csvPath = Path.Combine(baseDir, "summary.csv");

        // Ensure directories exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(sampleImageDir);
        Directory.CreateDirectory(extractedImageDir);

        // -------------------------------------------------
        // 1. Create deterministic sample images (PNG)
        // -------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(sampleImageDir, $"sample{i}.png");
            // Create a 100x100 bitmap with a distinct background color
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Simple solid color background
                graphics.Clear(i == 1 ? Color.Red : i == 2 ? Color.Green : Color.Blue);
                // Save the bitmap to a PNG file
                bitmap.Save(imagePath);
            }
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain images
        // -------------------------------------------------
        string[] sampleImages = Directory.GetFiles(sampleImageDir, "*.png");
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {docIndex} - contains sample images.");

            // Insert two images per document
            builder.InsertImage(sampleImages[(docIndex - 1) % sampleImages.Length]);
            builder.InsertParagraph();
            builder.InsertImage(sampleImages[(docIndex) % sampleImages.Length]);

            string docPath = Path.Combine(inputDir, $"SampleDocument{docIndex}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process DOCX files: extract images & CSV summary
        // -------------------------------------------------
        List<string> csvLines = new List<string>();
        csvLines.Add("Document,ImageFile,WidthPixels,HeightPixels,ImageType");

        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Retrieve image size information
                ImageSize imgSize = shape.ImageData.ImageSize;

                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build deterministic image file name
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(extractedImageDir, imageFileName);

                // Save the image to disk
                shape.ImageData.Save(imageFullPath);

                // Record CSV entry
                csvLines.Add($"{Path.GetFileName(docFile)},{imageFileName},{imgSize.WidthPixels},{imgSize.HeightPixels},{shape.ImageData.ImageType}");

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (csvLines.Count <= 1)
            throw new Exception("No images were extracted from the documents.");

        // Write CSV summary
        File.WriteAllLines(csvPath, csvLines);

        // Optional: indicate completion (no interactive input required)
        Console.WriteLine("Batch image extraction completed.");
        Console.WriteLine($"Extracted images folder: {extractedImageDir}");
        Console.WriteLine($"CSV summary file: {csvPath}");
    }
}
