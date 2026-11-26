using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json; // Required package as per rules

public class Program
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files containing the image
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"SampleDocument{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Batch process: extract images from each DOCX
        var extractedImagePaths = new System.Collections.Generic.List<string>();
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string ext = shape.ImageData.ImageType.ToString().ToLower(); // e.g., png, jpeg
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.{ext}";
                    string imagePath = Path.Combine(extractedDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create a zip archive containing all extracted images
        string zipPath = Path.Combine(outputDir, "ExtractedImages.zip");
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (string imgPath in extractedImagePaths)
            {
                archive.CreateEntryFromFile(imgPath, Path.GetFileName(imgPath));
            }
        }

        // Validate zip creation
        if (!File.Exists(zipPath) || new FileInfo(zipPath).Length == 0)
            throw new InvalidOperationException("Failed to create the zip archive with extracted images.");

        // Cleanup extracted images (optional)
        foreach (string imgPath in extractedImagePaths)
        {
            File.Delete(imgPath);
        }
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a white bitmap and draw a red rectangle
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (var pen = new Pen(Color.Red, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
