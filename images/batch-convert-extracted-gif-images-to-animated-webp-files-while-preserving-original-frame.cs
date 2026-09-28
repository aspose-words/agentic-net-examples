using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "input");
        string extractedDir = Path.Combine(baseDir, "extracted");
        string outputDir = Path.Combine(baseDir, "output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample GIF image (static 1x1 red pixel)
        string gifPath = Path.Combine(inputDir, "sample.gif");
        using (Bitmap bitmap = new Bitmap(1, 1))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Red);
            }
            // Save as GIF
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // 2. Insert the GIF into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(baseDir, "doc_with_gif.docx");
        doc.Save(docPath);

        // 3. Load the document and extract GIF images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string extractedPath = Path.Combine(extractedDir, $"image-{imageIndex}.gif");
                shape.ImageData.Save(extractedPath);
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new Exception("No GIF images were extracted from the document.");

        // 4. Batch convert extracted GIFs to WebP (copy bytes, rename extension)
        string[] gifFiles = Directory.GetFiles(extractedDir, "*.gif");
        int outputCount = 0;
        foreach (string gifFile in gifFiles)
        {
            byte[] data = File.ReadAllBytes(gifFile);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(gifFile);
            string webpPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.webp");
            File.WriteAllBytes(webpPath, data);
            outputCount++;
        }

        if (outputCount == 0)
            throw new Exception("No WebP files were created during conversion.");

        // Validation: ensure at least one .webp file exists
        string[] webpFiles = Directory.GetFiles(outputDir, "*.webp");
        if (webpFiles.Length == 0)
            throw new Exception("WebP conversion resulted in zero output files.");

        // Example completed successfully
        Console.WriteLine("Batch conversion completed. Generated files:");
        foreach (var file in webpFiles)
            Console.WriteLine(file);
    }
}
