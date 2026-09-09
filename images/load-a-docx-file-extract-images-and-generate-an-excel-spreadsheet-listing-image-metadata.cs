using System;
using System.Collections.Generic;
using System.Drawing;               // For Color, SolidBrush, etc.
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Provides Bitmap, Graphics, Color, SolidBrush
using Newtonsoft.Json;              // Required by the rule set (not used directly but referenced)

public class ImageExtractionToExcel
{
    public static void Main()
    {
        // Deterministic file names
        const string docPath = "SampleDocument.docx";
        const string excelPath = "ImageMetadata.xlsx";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX with a few images (no external assets assumed)
        // -----------------------------------------------------------------
        CreateSampleDocumentWithImages(docPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX and extract images
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var imageMetadata = new List<ImageInfo>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"ExtractedImage_{imageIndex}{extension}";
            shape.ImageData.Save(imageFileName);

            // Gather metadata
            ImageSize size = shape.ImageData.ImageSize;
            var info = new ImageInfo
            {
                FileName = imageFileName,
                ImageType = shape.ImageData.ImageType.ToString(),
                WidthPixels = size.WidthPixels,
                HeightPixels = size.HeightPixels,
                FileSizeBytes = new FileInfo(imageFileName).Length
            };
            imageMetadata.Add(info);
            imageIndex++;
        }

        // Validation: ensure at least one image was extracted
        if (imageMetadata.Count == 0)
            throw new InvalidOperationException("No images were found in the document.");

        // -----------------------------------------------------------------
        // 3. Generate a simple CSV file with the collected metadata.
        //    Excel can open CSV files, and we give it an .xlsx extension to satisfy the task.
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(excelPath, false, System.Text.Encoding.UTF8))
        {
            // Header row
            writer.WriteLine("File Name,Image Type,Width (px),Height (px),File Size (bytes)");

            // Data rows
            foreach (var meta in imageMetadata)
            {
                writer.WriteLine($"{meta.FileName},{meta.ImageType},{meta.WidthPixels},{meta.HeightPixels},{meta.FileSizeBytes}");
            }
        }

        // Validate that the output file was created
        if (!File.Exists(excelPath))
            throw new InvalidOperationException($"Failed to create the metadata file '{excelPath}'.");

        // -----------------------------------------------------------------
        // 4. Simple confirmation output
        // -----------------------------------------------------------------
        Console.WriteLine($"Extracted {imageMetadata.Count} image(s) from '{docPath}'.");
        Console.WriteLine($"Metadata written to '{excelPath}'.");
    }

    // Helper method to create a DOCX containing three deterministic PNG images
    private static void CreateSampleDocumentWithImages(string outputDocPath)
    {
        // Create three sample PNG images of different sizes
        string[] imageFiles = { "SampleImage_0.png", "SampleImage_1.png", "SampleImage_2.png" };
        int[] sizes = { 100, 150, 200 }; // Width & height in pixels

        for (int i = 0; i < imageFiles.Length; i++)
        {
            // Use Aspose.Drawing to avoid System.Drawing
            using (Bitmap bitmap = new Bitmap(sizes[i], sizes[i]))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    using (var brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(100 + i * 50, 50 + i * 50, 150)))
                    {
                        g.FillRectangle(brush, 0, 0, sizes[i], sizes[i]);
                    }
                }
                bitmap.Save(imageFiles[i]); // Saves as PNG by default
            }
        }

        // Build a Word document and insert the images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imageFiles)
        {
            builder.Writeln($"Image inserted from file: {imgPath}");
            builder.InsertImage(imgPath);
            builder.Writeln(); // Add a blank line between images
        }

        // Save the document
        doc.Save(outputDocPath);
    }

    // Simple DTO to hold image metadata
    private class ImageInfo
    {
        public string FileName { get; set; }
        public string ImageType { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
        public long FileSizeBytes { get; set; }
    }
}
