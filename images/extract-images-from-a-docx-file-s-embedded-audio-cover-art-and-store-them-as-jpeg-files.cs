using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;                 // Aspose.Drawing.Common provides Bitmap, Graphics, Color
using Aspose.Drawing.Imaging;        // Provides ImageFormat for saving images

public class Program
{
    public static void Main()
    {
        // Define folders for artifacts.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image that will act as the audio cover art.
        // -----------------------------------------------------------------
        string coverImagePath = Path.Combine(artifactsDir, "cover.jpg");
        CreateSampleJpeg(coverImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Build a DOCX document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, coverImagePath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images (cover art) as JPEG files.
        // -----------------------------------------------------------------
        ExtractImagesAsJpeg(docPath, artifactsDir);
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        // Create a bitmap and draw a simple rectangle.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a blue rectangle.
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
            {
                graphics.FillRectangle(brush, 20, 20, width - 40, height - 40);
            }
            // Save as JPEG using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Validate that the image file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image inline. InsertImage already adds the shape to the document,
        // so no additional AppendChild is required.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        // Validate that the document file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create document at '{docPath}'.");
    }

    // Loads a DOCX file, extracts all images, and saves them as JPEG files.
    private static void ExtractImagesAsJpeg(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Force JPEG extension if the original is not JPEG.
                if (!extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase))
                    extension = ".jpeg";

                string outputPath = Path.Combine(outputDir, $"extracted_{imageIndex}{extension}");
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
