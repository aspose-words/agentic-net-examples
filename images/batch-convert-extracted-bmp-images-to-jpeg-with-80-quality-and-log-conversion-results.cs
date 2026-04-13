using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageConversion");
        string inputDir = Path.Combine(baseDir, "Input");
        string extractedDir = Path.Combine(baseDir, "Extracted");
        string outputDir = Path.Combine(baseDir, "Converted");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create sample BMP images
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 100))
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                Aspose.Drawing.Color rectColor = i == 1 ? Aspose.Drawing.Color.Red :
                                                i == 2 ? Aspose.Drawing.Color.Green :
                                                         Aspose.Drawing.Color.Blue;
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(rectColor))
                {
                    g.FillRectangle(brush, 10, 10, 180, 80);
                }
                bitmap.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
            }
        }

        // 2. Insert BMP images into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpFile in Directory.GetFiles(inputDir, "*.bmp"))
        {
            builder.InsertImage(bmpFile);
            builder.Writeln(); // separate images
        }

        // 3. Extract images from the document
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedPath = Path.Combine(extractedDir, $"extracted_{extractedCount}{extension}");
                shape.ImageData.Save(extractedPath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 4. Batch convert extracted images to JPEG with 80% quality
        int convertedCount = 0;
        foreach (string imgPath in Directory.GetFiles(extractedDir))
        {
            // Load the extracted image into a temporary document
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            tempBuilder.InsertImage(imgPath);

            // Configure JPEG save options with 80% quality
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80
            };

            string jpegPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(imgPath) + ".jpg");
            tempDoc.Save(jpegPath, jpegOptions);

            Console.WriteLine($"Converted '{Path.GetFileName(imgPath)}' to '{Path.GetFileName(jpegPath)}' with 80% quality.");
            convertedCount++;
        }

        if (convertedCount == 0)
            throw new InvalidOperationException("No JPEG images were produced during conversion.");

        Console.WriteLine($"Batch conversion completed: {extractedCount} image(s) extracted, {convertedCount} JPEG(s) created.");
    }
}
