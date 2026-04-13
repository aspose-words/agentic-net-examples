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
        // Create a deterministic JPEG image to be used in the document.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create a Word document and insert the JPEG image.
        const string sourceDocPath = "source.docx";
        CreateDocumentWithImage(sourceDocPath, sampleImagePath);

        // Load the document and process all JPEG images.
        Document doc = new Document(sourceDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int savedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Apply grayscale rendering to the shape's image data.
            shape.ImageData.GrayScale = true;

            // Determine output file name with proper extension.
            string outFileName = $"ExtractedImage_{savedCount}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";

            // Save the (now grayscale) image to the file system.
            shape.ImageData.Save(outFileName);
            savedCount++;
        }

        // Validate that at least one image was saved.
        if (savedCount == 0)
            throw new InvalidOperationException("No JPEG images were found and saved.");

        // Clean up the temporary files created for the example.
        CleanupTemporaryFiles(sampleImagePath, sourceDocPath);
    }

    private static void CreateSampleJpeg(string filePath)
    {
        // Create a 100x100 bitmap, fill with white, draw a red ellipse, and save as JPEG.
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                graphics.FillEllipse(brush, 10, 10, 80, 80);
            }

            // Save the bitmap as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void CleanupTemporaryFiles(params string[] files)
    {
        foreach (string file in files)
        {
            try
            {
                if (File.Exists(file))
                    File.Delete(file);
            }
            catch
            {
                // Ignored – cleanup is best‑effort.
            }
        }
    }
}
