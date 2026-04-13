using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define working directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string docsDir = Path.Combine(baseDir, "Docs");
        string zipPath = Path.Combine(baseDir, "ExtractedImages.zip");

        // Ensure clean environment
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);

        // 1. Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(docsDir, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(docsDir, "Doc2.docx"), sampleImagePath);

        // 3. Extract images from each DOCX file
        int totalExtracted = 0;
        foreach (string docPath in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    if (!File.Exists(imageFullPath))
                        throw new InvalidOperationException($"Failed to save extracted image: {imageFullPath}");
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // 4. Create a zip archive containing all extracted images
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(imagesDir, zipPath);
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the zip archive.");

        // Output paths for verification (optional)
        Console.WriteLine($"Extracted {totalExtracted} images to: {imagesDir}");
        Console.WriteLine($"Created zip archive: {zipPath}");
    }

    // Creates a simple white background PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
            {
                g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Creates a DOCX file with a single image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image inline
        builder.InsertImage(imagePath);
        // Add a paragraph to separate possible multiple images
        builder.Writeln();
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document: {docPath}");
    }
}
