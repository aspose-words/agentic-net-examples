using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputRoot = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputRoot);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a few ODT documents containing the sample image
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.odt");
            CreateOdtWithImages(docPath, sampleImagePath, i);
        }

        // Batch extract images from each ODT file
        foreach (string odtFile in Directory.GetFiles(inputFolder, "*.odt"))
        {
            string docName = Path.GetFileNameWithoutExtension(odtFile);
            string docOutputFolder = Path.Combine(outputRoot, docName);
            Directory.CreateDirectory(docOutputFolder);

            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string extension = GetImageExtension(shape.ImageData.ImageType);
                    string outPath = Path.Combine(docOutputFolder, $"image-{imageIndex}{extension}");
                    shape.ImageData.Save(outPath);
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{odtFile}'.");
        }
    }

    // Creates a simple white PNG image with a black rectangle for deterministic content
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates an ODT document and inserts the sample image a number of times
    private static void CreateOdtWithImages(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageCount; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
        }

        doc.Save(docPath);
    }

    // Maps Aspose.Words.Drawing.ImageType to a file extension
    private static string GetImageExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpg";
            case ImageType.Png:
                return ".png";
            case ImageType.Gif:
                return ".gif";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            default:
                return ".img";
        }
    }
}
