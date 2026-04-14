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
        // Base output directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string odtDir = Path.Combine(baseDir, "OdtFiles");
        string catalogPath = Path.Combine(baseDir, "ImageCatalog.pdf");

        // Ensure clean folders.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(odtDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        const int imgWidth = 200;
        const int imgHeight = 200;
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Draw a simple rectangle.
        graphics.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
        // Save the image and release resources.
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a few ODT documents that contain the sample image.
        // -----------------------------------------------------------------
        const int odtCount = 3;
        for (int i = 1; i <= odtCount; i++)
        {
            Document odtDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(odtDoc);
            builder.Writeln($"Sample ODT document #{i}");
            // Insert the sample image twice to have multiple images per file.
            builder.InsertImage(sampleImagePath);
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);

            string odtPath = Path.Combine(odtDir, $"Sample{i}.odt");
            // Save as ODT using OdtSaveOptions.
            odtDoc.Save(odtPath, new OdtSaveOptions());
        }

        // -----------------------------------------------------------------
        // 3. Batch extract images from all ODT files.
        // -----------------------------------------------------------------
        string[] odtFiles = Directory.GetFiles(odtDir, "*.odt");
        foreach (string odtFile in odtFiles)
        {
            Document srcDoc = new Document(odtFile);
            NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new Exception($"No images were extracted from '{odtFile}'.");
        }

        // -----------------------------------------------------------------
        // 4. Create a searchable PDF catalog that lists all extracted images.
        // -----------------------------------------------------------------
        Document catalogDoc = new Document();
        DocumentBuilder catBuilder = new DocumentBuilder(catalogDoc);
        catBuilder.Writeln("Image Catalog");
        catBuilder.Font.Size = 16;
        catBuilder.Font.Bold = true;
        catBuilder.InsertParagraph();

        string[] extractedImages = Directory.GetFiles(imagesDir);
        foreach (string imgPath in extractedImages.OrderBy(p => p))
        {
            string fileName = Path.GetFileName(imgPath);
            catBuilder.Writeln(fileName);
            catBuilder.InsertImage(imgPath);
            catBuilder.InsertParagraph();
            // Optional page break after each image for readability.
            catBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Save the catalog as PDF.
        catalogDoc.Save(catalogPath, SaveFormat.Pdf);

        // Validate that the PDF catalog was created.
        if (!File.Exists(catalogPath))
            throw new Exception("Failed to create the PDF catalog.");

        // The program finishes without awaiting user input.
    }
}
