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
        // Paths
        string inputDocPath = "sample.docx";
        string sampleImagePath = "sample.png";
        string outputFolder = "ExtractedImages";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Create a deterministic sample image (100x100 white)
        // -------------------------------------------------
        int width = 100;
        int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // Create a sample DOCX containing the image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // Load the DOCX and extract all embedded images
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extension = GetExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}{extension}");
                shape.ImageData.Save(outputPath);
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException($"Failed to save extracted image to '{outputPath}'.");
                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Helper to map ImageType to file extension
    private static string GetExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpeg";
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
