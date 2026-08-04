using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string imagesDir = Path.Combine(baseDir, "Images");
        string docsDir = Path.Combine(baseDir, "Docs");
        string extractedDir = Path.Combine(baseDir, "Extracted");
        string zipPath = Path.Combine(baseDir, "ExtractedImages.zip");

        // Ensure clean environment.
        foreach (string dir in new[] { imagesDir, docsDir, extractedDir })
            Directory.CreateDirectory(dir);
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle.
                g.FillRectangle(Brushes.Blue, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a few DOCX files that contain the sample image.
        // -------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} with an embedded image.");
            // Insert the previously created image.
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(docsDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images.
        // -------------------------------------------------
        int extractedImageCount = 0;
        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outputFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image_{imageIndex}{extension}";
                    string outputPath = Path.Combine(extractedDir, outputFileName);
                    shape.ImageData.Save(outputPath);
                    extractedImageCount++;
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageCount == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 4. Create a ZIP archive containing all extracted images.
        // -------------------------------------------------
        ZipFile.CreateFromDirectory(extractedDir, zipPath);

        // Optional: indicate completion (no interactive I/O required).
        // The program ends here.
    }
}
