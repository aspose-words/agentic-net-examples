using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageExtractionToZip
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "Images.zip");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);

        // Sample PNG image (1x1 pixel) in Base64
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Create a few sample documents each containing an image shape
        for (int docIndex = 0; docIndex < 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imgStream);
            }
            string docPath = Path.Combine(inputFolder, $"doc{docIndex}.docx");
            doc.Save(docPath);
        }

        // Collect all image files extracted from the documents
        List<string> extractedImagePaths = new List<string>();

        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document loadedDoc = new Document(docFile);
            // Find all shape nodes that contain images
            var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                      .OfType<Shape>()
                                      .Where(s => s.HasImage);
            int imgIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imgIndex}.png";
                string imagePath = Path.Combine(imageFolder, imageFileName);
                shape.ImageData.Save(imagePath);
                extractedImagePaths.Add(imagePath);
                imgIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document collection.");

        // Create a ZIP archive containing all extracted images
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string imgPath in extractedImagePaths)
            {
                string entryName = Path.GetFileName(imgPath);
                archive.CreateEntryFromFile(imgPath, entryName);
            }
        }

        // Verify ZIP creation
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        // Optional cleanup of temporary image files (keep ZIP)
        foreach (string imgPath in extractedImagePaths)
        {
            if (File.Exists(imgPath))
                File.Delete(imgPath);
        }

        // Program ends without interactive prompts
    }
}
