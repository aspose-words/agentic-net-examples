using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image larger than 800px width.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 1200, 800); // 1200x800 pixels

        // 2. Build a Word document that contains the JPEG image twice.
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImages(inputDocPath, sampleJpegPath);

        // 3. Load the document and resize all JPEG images to a maximum width of 800px.
        Document doc = new Document(inputDocPath);
        bool anyJpegProcessed = false;

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Load the image bytes into an Aspose.Drawing.Image.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream srcStream = new MemoryStream(imageBytes))
            using (Aspose.Drawing.Image srcImage = Aspose.Drawing.Image.FromStream(srcStream))
            {
                int originalWidth = srcImage.Width;
                int originalHeight = srcImage.Height;

                // If the width is already <= 800, no resizing needed.
                if (originalWidth <= 800)
                    continue;

                // Calculate new dimensions while preserving aspect ratio.
                int newWidth = 800;
                int newHeight = (int)Math.Round((double)originalHeight * newWidth / originalWidth);

                // Create a new bitmap with the target size.
                using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(newWidth, newHeight))
                using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                {
                    // High‑quality scaling.
                    graphics.InterpolationMode = Aspose.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(srcImage, 0, 0, newWidth, newHeight);

                    // Save the resized bitmap to a memory stream in JPEG format.
                    using (MemoryStream resizedStream = new MemoryStream())
                    {
                        resizedBitmap.Save(resizedStream, ImageFormat.Jpeg);
                        resizedStream.Position = 0; // Reset before reuse.

                        // Replace the image in the shape with the resized version.
                        shape.ImageData.SetImage(resizedStream);
                    }
                }
            }

            anyJpegProcessed = true;
        }

        if (!anyJpegProcessed)
            throw new InvalidOperationException("No JPEG images were found to process.");

        // 4. Save the modified document.
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath);

        // Optional: Export the resized images to verify the result.
        ExportJpegImages(doc, artifactsDir);
    }

    // Creates a solid‑color JPEG image of specified dimensions.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Builds a simple document and inserts the given image twice.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with JPEG images:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }

    // Saves all JPEG images from the document to the artifacts folder for verification.
    private static void ExportJpegImages(Document doc, string folder)
    {
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                string outPath = Path.Combine(folder, $"extracted_{index}.jpg");
                shape.ImageData.Save(outPath);
                index++;
            }
        }

        if (index == 0)
            throw new InvalidOperationException("No JPEG images were extracted for verification.");
    }
}
