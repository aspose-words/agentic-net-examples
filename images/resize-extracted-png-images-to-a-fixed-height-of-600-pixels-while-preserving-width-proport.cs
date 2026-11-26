using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath, 200, 200);

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, inputImagePath);

        // Load the document and process PNG images.
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int resizedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream inputStream = new MemoryStream(imageBytes))
            {
                inputStream.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(inputStream))
                {
                    // Calculate new dimensions preserving aspect ratio with height = 600.
                    const int targetHeight = 600;
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;
                    int targetWidth = (int)Math.Round(originalWidth * (targetHeight / (double)originalHeight));

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.White);
                            graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                        }

                        // Save the resized image.
                        string resizedPath = $"resized-{resizedCount}.png";
                        resizedBitmap.Save(resizedPath);
                        resizedCount++;

                        // Optional: replace the image in the document with the resized version.
                        shape.ImageData.SetImage(resizedPath);
                    }
                }
            }
        }

        // Validate that at least one image was resized.
        if (resizedCount == 0)
            throw new InvalidOperationException("No PNG images were found and resized.");

        // Save the updated document.
        const string outputDocPath = "sample_resized.docx";
        doc.Save(outputDocPath);
    }

    private static void CreateSamplePng(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
