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
        // Base directories for input documents, sample images and extracted images.
        string baseDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(baseDir, "SampleImages");
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure clean environment.
        if (Directory.Exists(imagesDir)) Directory.Delete(imagesDir, true);
        if (Directory.Exists(docsDir)) Directory.Delete(docsDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic sample images of different formats.
        CreateSampleImage(Path.Combine(imagesDir, "sample.png"), 100, 100, Aspose.Drawing.Color.LightBlue, ImageFormat.Png);
        CreateSampleImage(Path.Combine(imagesDir, "sample.jpg"), 120, 80, Aspose.Drawing.Color.LightCoral, ImageFormat.Jpeg);
        CreateSampleImage(Path.Combine(imagesDir, "sample.bmp"), 80, 120, Aspose.Drawing.Color.LightGreen, ImageFormat.Bmp);
        CreateSampleImage(Path.Combine(imagesDir, "sample.gif"), 90, 90, Aspose.Drawing.Color.LightYellow, ImageFormat.Gif);

        // Build a sample document that contains the images.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Document with multiple image formats:");
        builder.InsertImage(Path.Combine(imagesDir, "sample.png"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample.jpg"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample.bmp"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesDir, "sample.gif"));
        string sampleDocPath = Path.Combine(docsDir, "SampleDocument.docx");
        sampleDoc.Save(sampleDocPath, SaveFormat.Docx);

        // Batch process all DOC/DOCX files in the input folder.
        var docFiles = Directory.GetFiles(docsDir, "*.*", SearchOption.TopDirectoryOnly)
                                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                                .ToArray();

        int totalExtracted = 0;

        foreach (var docFile in docFiles)
        {
            Document doc = new Document(docFile);
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;
            foreach (var shape in shapes)
            {
                // Determine image type and corresponding file extension.
                ImageType imgType = shape.ImageData.ImageType;
                string extension = FileFormatUtil.ImageTypeToExtension(imgType); // includes leading dot.
                string formatFolderName = extension.TrimStart('.').ToLowerInvariant();

                // Create subfolder for this image format.
                string formatFolderPath = Path.Combine(outputDir, formatFolderName);
                Directory.CreateDirectory(formatFolderPath);

                // Build deterministic file name.
                string docName = Path.GetFileNameWithoutExtension(docFile);
                string imageFileName = $"{docName}_img{imageIndex}{extension}";
                string imagePath = Path.Combine(formatFolderPath, imageFileName);

                // Save the image.
                shape.ImageData.Save(imagePath);
                imageIndex++;
                totalExtracted++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Optional: verify that each format subfolder contains files.
        var formatFolders = Directory.GetDirectories(outputDir);
        foreach (var folder in formatFolders)
        {
            if (!Directory.GetFiles(folder).Any())
                throw new InvalidOperationException($"Expected images in folder '{folder}' but none were found.");
        }

        // The program finishes without interactive prompts.
    }

    // Helper method to create a deterministic bitmap and save it.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, ImageFormat format)
    {
        // Create a bitmap with the requested size.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a graphics object to draw onto the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with the specified color.
                graphics.Clear(backColor);
            }

            // Save the bitmap to the specified file using the requested image format.
            bitmap.Save(filePath, format);
        }
    }
}
