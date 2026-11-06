using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // Create a deterministic sample image
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), sampleImagePath);

        // Extract images and build index
        var indexEntries = new List<(string DocumentPath, string ImagePath)>();
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.png";
                    string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    indexEntries.Add((docPath, imageFullPath));
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (indexEntries.Count == 0)
            throw new Exception("No images were extracted from the documents.");

        // Generate CSV index that Excel can open
        string csvPath = Path.Combine(baseDir, "ImageIndex.csv");
        using (var writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("DocumentPath,ImagePath");
            foreach (var entry in indexEntries)
            {
                // Quote paths in case they contain commas
                writer.WriteLine($"\"{entry.DocumentPath}\",\"{entry.ImagePath}\"");
            }
        }

        // Simple validation that CSV was created
        if (!File.Exists(csvPath))
            throw new Exception("Failed to create the CSV index file.");
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 200x200 white bitmap
        Bitmap bitmap = new Bitmap(200, 200);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Save as PNG
        bitmap.Save(filePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the previously created image
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
