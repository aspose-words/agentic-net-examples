using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class GrayscaleImageExtractor
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG image.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 200, 100);

        // 2. Build a DOCX that contains several copies of the JPEG image.
        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        CreateDocumentWithImages(sourceDocPath, sampleJpegPath, 3);

        // 3. Load the document and extract JPEG images, applying a grayscale filter.
        Document doc = new Document(sourceDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int jpegIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Apply grayscale rendering flag.
            shape.ImageData.GrayScale = true;

            // Determine file name with proper extension.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string grayImagePath = Path.Combine(artifactsDir, $"extracted_{jpegIndex}_gray{extension}");

            // Save the grayscale image.
            shape.ImageData.Save(grayImagePath);
            jpegIndex++;
        }

        // Validation.
        if (jpegIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and processed.");

        Console.WriteLine($"Processed {jpegIndex} JPEG image(s). Grayscale files are located in: {artifactsDir}");
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }
            }
            // Explicitly save as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a DOCX file and inserts the specified image multiple times.
    private static void CreateDocumentWithImages(string docPath, string imagePath, int repeatCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < repeatCount; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
        }

        doc.Save(docPath);
    }
}
