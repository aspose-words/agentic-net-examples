using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string extractedFolder = "ExtractedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);

        // Create a deterministic sample image
        string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files containing the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // List to hold paths of extracted images
        List<string> extractedImagePaths = new List<string>();

        // Batch process each DOCX file
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string extension = shape.ImageData.ImageType.ToString().ToLower(); // e.g., png, jpeg
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.{extension}";
                    string imageFilePath = Path.Combine(extractedFolder, imageFileName);
                    shape.ImageData.Save(imageFilePath);
                    extractedImagePaths.Add(imageFilePath);
                    imageIndex++;
                }
            }
        }

        // Validate that images were extracted
        if (extractedImagePaths.Count == 0)
            throw new Exception("No images were extracted from the documents.");

        // Create a zip archive containing all extracted images
        string zipPath = "ExtractedImages.zip";
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (string imagePath in extractedImagePaths)
            {
                archive.CreateEntryFromFile(imagePath, Path.GetFileName(imagePath));
            }
        }

        // Validate zip creation
        if (!File.Exists(zipPath) || new FileInfo(zipPath).Length == 0)
            throw new Exception("Failed to create the zip archive of extracted images.");
    }

    private static void CreateSampleImage(string path)
    {
        int width = 100;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Black))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path);
        }
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
