using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Newtonsoft.Json;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string csvPath = Path.Combine(baseDir, "summary.csv");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // Create deterministic sample images.
        CreateSampleImage(Path.Combine(baseDir, "sample1.png"), 200, 200, Color.LightBlue);
        CreateSampleImage(Path.Combine(baseDir, "sample2.png"), 300, 150, Color.LightGreen);
        CreateSampleImage(Path.Combine(baseDir, "sample3.png"), 100, 250, Color.LightCoral);

        // Create sample DOCX files that contain the images.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"),
            new[] { Path.Combine(baseDir, "sample1.png"), Path.Combine(baseDir, "sample2.png") });

        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"),
            new[] { Path.Combine(baseDir, "sample2.png"), Path.Combine(baseDir, "sample3.png") });

        CreateSampleDocument(Path.Combine(inputDir, "Sample3.docx"),
            new[] { Path.Combine(baseDir, "sample1.png"), Path.Combine(baseDir, "sample3.png") });

        // Prepare CSV header.
        var csvLines = new List<string>
        {
            "DocumentName,ImageIndex,ImageFileName,WidthPixels,HeightPixels,ImageType"
        };

        // Process each DOCX file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imageOutputDir, imageFileName);

                // Save the image to the output folder.
                shape.ImageData.Save(imageFullPath);

                // Retrieve image size information.
                ImageSize size = shape.ImageData.ImageSize;

                // Add a line to the CSV summary.
                csvLines.Add($"{Path.GetFileName(docPath)},{imageIndex},{imageFileName},{size.WidthPixels},{size.HeightPixels},{shape.ImageData.ImageType}");

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (csvLines.Count <= 1)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Write the CSV summary file.
        File.WriteAllLines(csvPath, csvLines);

        // Optional: output locations for verification.
        Console.WriteLine($"Images extracted to: {imageOutputDir}");
        Console.WriteLine($"CSV summary created at: {csvPath}");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file and inserts the specified images.
    private static void CreateSampleDocument(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imagePaths)
        {
            // Ensure the image file exists before insertion.
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");

            // Insert the image inline.
            Shape shape = builder.InsertImage(imgPath);
            // Append a line break after each image for readability.
            builder.Writeln();
        }

        // Save the document.
        doc.Save(docPath);
    }
}
