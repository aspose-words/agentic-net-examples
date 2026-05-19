using System;
using System.IO;
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
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic sample images.
        string pngPath = Path.Combine(baseDir, "sample1.png");
        string jpgPath = Path.Combine(baseDir, "sample2.jpg");
        CreateSampleImage(pngPath, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(jpgPath, 250, 180, Aspose.Drawing.Color.LightCoral);

        // Create sample Word documents that contain the images.
        CreateWordDocument(Path.Combine(inputDir, "Doc1.docx"), pngPath);
        CreateWordDocument(Path.Combine(inputDir, "Doc2.docx"), jpgPath);

        // Batch convert extracted images to WebP.
        int outputCount = 0;
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Save the original image to a temporary file.
                string tempImgPath = Path.Combine(outputDir,
                    $"temp_{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                shape.ImageData.Save(tempImgPath);

                // Create a new document that contains only this image.
                Document imgDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(imgDoc);
                builder.InsertImage(tempImgPath);

                // Define the output WebP file name.
                string outFile = Path.Combine(
                    outputDir,
                    $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.webp");

                // Save the document page (which holds the image) as WebP.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.WebP);
                imgDoc.Save(outFile, options);
                outputCount++;

                // Clean up the temporary image file.
                if (File.Exists(tempImgPath))
                    File.Delete(tempImgPath);

                imageIndex++;
            }
        }

        // Validate that at least one WebP file was produced.
        if (outputCount == 0)
            throw new InvalidOperationException("No images were converted to WebP.");

        // List output files.
        foreach (string file in Directory.GetFiles(outputDir, "*.webp"))
        {
            Console.WriteLine("Created: " + file);
        }
    }

    // Helper to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }

    // Helper to create a Word document with a single image.
    private static void CreateWordDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
