using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json; // Required by the task, even if not used

public class Program
{
    public static void Main()
    {
        // Base directories for the example data.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExampleData");
        string docsDir = Path.Combine(baseDir, "Docs");
        string imagesDir = Path.Combine(baseDir, "SourceImages");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images of different formats.
        // -----------------------------------------------------------------
        CreateSampleImage(Path.Combine(imagesDir, "sample_png.png"), 100, 100, Aspose.Drawing.Color.Red, ImageFormat.Png);
        CreateSampleImage(Path.Combine(imagesDir, "sample_jpeg.jpg"), 120, 80, Aspose.Drawing.Color.Blue, ImageFormat.Jpeg);
        CreateSampleImage(Path.Combine(imagesDir, "sample_bmp.bmp"), 80, 80, Aspose.Drawing.Color.Green, ImageFormat.Bmp);
        CreateSampleImage(Path.Combine(imagesDir, "sample_gif.gif"), 60, 60, Aspose.Drawing.Color.Yellow, ImageFormat.Gif);

        // -----------------------------------------------------------------
        // 2. Create a sample DOCX that contains the images.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(docsDir, "SampleDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with multiple image formats:");
        builder.InsertImage(Path.Combine(imagesDir, "sample_png.png"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample_jpeg.jpg"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample_bmp.bmp"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample_gif.gif"));

        doc.Save(docPath);

        // Create a second copy to demonstrate batch processing.
        string docPath2 = Path.Combine(docsDir, "SampleDocument2.docx");
        doc.Save(docPath2);

        // -----------------------------------------------------------------
        // 3. Batch extract images from all DOC/DOCX files.
        // -----------------------------------------------------------------
        string[] docFiles = Directory.GetFiles(docsDir, "*.*", SearchOption.TopDirectoryOnly)
                                     .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                 f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                                     .ToArray();

        if (docFiles.Length == 0)
            throw new InvalidOperationException("No document files found for processing.");

        foreach (string file in docFiles)
        {
            Document loadDoc = new Document(file);
            NodeCollection shapeNodes = loadDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                ImageType imgType = shape.ImageData.ImageType;
                string extension = FileFormatUtil.ImageTypeToExtension(imgType); // includes leading dot
                string formatFolder = Path.Combine(outputDir, extension.TrimStart('.').ToLowerInvariant());
                Directory.CreateDirectory(formatFolder);

                string imageFileName = $"{Path.GetFileNameWithoutExtension(file)}_img_{imageIndex}{extension}";
                string fullPath = Path.Combine(formatFolder, imageFileName);
                shape.ImageData.Save(fullPath);
                imageIndex++;
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{Path.GetFileName(file)}'.");
        }

        // -----------------------------------------------------------------
        // 4. Simple verification output.
        // -----------------------------------------------------------------
        Console.WriteLine("Image extraction completed. Extracted files:");
        foreach (string dir in Directory.GetDirectories(outputDir))
        {
            string format = new DirectoryInfo(dir).Name;
            string[] files = Directory.GetFiles(dir);
            Console.WriteLine($"  Format '{format}': {files.Length} file(s)");
            foreach (string f in files)
                Console.WriteLine($"    {Path.GetFileName(f)}");
        }
    }

    // Helper method to create a solid‑color image and save it in the requested format.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor, ImageFormat format)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(fillColor);
            bitmap.Save(filePath, format);
        }
    }
}
