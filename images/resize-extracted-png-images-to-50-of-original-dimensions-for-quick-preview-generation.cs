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
        // Create a deterministic PNG image to be used as input.
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath, 200, 200);

        // Create a Word document and insert the PNG image twice.
        const string docPath = "document.docx";
        CreateDocumentWithImages(docPath, inputImagePath);

        // Load the document and extract PNG images, resizing each to 50% of its original size.
        const string outputFolder = "previews";
        Directory.CreateDirectory(outputFolder);
        ExtractAndResizePngImages(docPath, outputFolder);
    }

    // Creates a simple PNG image with a red rectangle on a white background.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Brush brush = new SolidBrush(Color.Red))
            {
                graphics.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
            }
            bitmap.Save(filePath);
        }

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image: {filePath}");
    }

    // Creates a Word document and inserts the specified image file.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image twice to demonstrate handling multiple occurrences.
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);

        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document: {docPath}");
    }

    // Extracts PNG images from the document, resizes them to 50%, and saves the previews.
    private static void ExtractAndResizePngImages(string docPath, string outputFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image using Aspose.Drawing.
                using (Image originalImage = Image.FromStream(originalStream))
                {
                    int originalWidth = originalImage.Width;
                    int originalHeight = originalImage.Height;

                    // Calculate new dimensions (50% of original).
                    int newWidth = Math.Max(1, originalWidth / 2);
                    int newHeight = Math.Max(1, originalHeight / 2);

                    // Create a new bitmap with the reduced size.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(
                            originalImage,
                            new Rectangle(0, 0, newWidth, newHeight),
                            new Rectangle(0, 0, originalWidth, originalHeight),
                            GraphicsUnit.Pixel);

                        string previewPath = Path.Combine(outputFolder, $"preview_{imageIndex}.png");
                        resizedBitmap.Save(previewPath);

                        if (!File.Exists(previewPath))
                            throw new Exception($"Failed to save resized preview: {previewPath}");
                    }
                }
            }

            imageIndex++;
        }

        if (imageIndex == 0)
            throw new Exception("No PNG images were found and resized in the document.");
    }
}
