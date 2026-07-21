using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string summaryCsvPath = Path.Combine(baseDir, "ImageSummary.csv");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // Create deterministic sample images
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 100, Color.LightBlue);
        CreateSampleImage(sampleImage2, 150, 150, Color.LightGreen);

        // Create sample DOCX files that contain the images
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), new[] { sampleImage1, sampleImage2 });
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), new[] { sampleImage2 });

        // Batch process all DOCX files in the input folder
        var csvLines = new List<string>();
        csvLines.Add("Document,ImageFile,WidthPixels,HeightPixels,Extension");

        int totalImages = 0;

        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue; // Skip shapes without images

                ImageData imgData = shape.ImageData;
                string extension = FileFormatUtil.ImageTypeToExtension(imgData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imageOutputDir, imageFileName);

                // Save the extracted image
                imgData.Save(imageFullPath);

                // Retrieve image size in pixels
                int widthPx = imgData.ImageSize.WidthPixels;
                int heightPx = imgData.ImageSize.HeightPixels;

                // Add a line to the CSV summary
                csvLines.Add($"{Path.GetFileName(docPath)},{imageFileName},{widthPx},{heightPx},{extension}");

                imageIndex++;
                totalImages++;
            }
        }

        // Write CSV summary
        File.WriteAllLines(summaryCsvPath, csvLines);

        // Simple validation
        if (totalImages == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX document and inserts the provided images
    private static void CreateSampleDocument(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imagePaths)
        {
            // Insert each image on a new paragraph
            builder.Writeln();
            builder.InsertImage(imgPath);
        }

        doc.Save(docPath);
    }
}
