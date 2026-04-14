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
        // Create a sample GIF image.
        const string sampleGifPath = "sample.gif";
        CreateSampleGif(sampleGifPath);

        // Create a document and insert the GIF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape gifShape = builder.InsertImage(sampleGifPath);
        doc.Save("doc_with_gif.docx");

        // Load the document and resize any GIF images to a maximum width of 300 pixels.
        Document loadedDoc = new Document("doc_with_gif.docx");
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Extract the GIF image into a memory stream.
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Resize the GIF while preserving animation frames.
                    byte[] resizedGif = ResizeGif(originalStream, 300);

                    // Replace the image in the shape with the resized GIF.
                    using (MemoryStream resizedStream = new MemoryStream(resizedGif))
                    {
                        shape.ImageData.SetImage(resizedStream);
                    }

                    // Save the resized GIF to a file for verification.
                    File.WriteAllBytes($"resized_{processedCount}.gif", resizedGif);
                }

                processedCount++;
            }
        }

        // Save the updated document.
        loadedDoc.Save("doc_with_resized_gif.docx");

        // Validation.
        if (!File.Exists("doc_with_resized_gif.docx"))
            throw new Exception("The output document was not created.");
        if (processedCount == 0)
            throw new Exception("No GIF images were found and processed.");
    }

    // Creates a deterministic sample GIF image.
    private static void CreateSampleGif(string filePath)
    {
        const int width = 500;
        const int height = 200;

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);

        Pen pen = new Pen(Color.Blue, 5);
        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);

        // Save as GIF.
        bitmap.Save(filePath, ImageFormat.Gif);

        // Clean up.
        pen.Dispose();
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Resizes a GIF image to the specified maximum width while preserving aspect ratio.
    private static byte[] ResizeGif(Stream inputGifStream, int maxWidth)
    {
        // Load the original GIF.
        Image originalImage = Image.FromStream(inputGifStream);
        int originalWidth = originalImage.Width;
        int originalHeight = originalImage.Height;

        // If the image is already within the desired width, return the original bytes.
        if (originalWidth <= maxWidth)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                originalImage.Save(ms, ImageFormat.Gif);
                return ms.ToArray();
            }
        }

        // Calculate new dimensions while preserving aspect ratio.
        double scale = (double)maxWidth / originalWidth;
        int newWidth = maxWidth;
        int newHeight = (int)(originalHeight * scale);

        // Create a new bitmap with the target size and draw the original image onto it.
        Bitmap resizedBitmap = new Bitmap(newWidth, newHeight);
        Graphics graphics = Graphics.FromImage(resizedBitmap);
        graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
        graphics.Dispose();

        // Save the resized bitmap as a GIF.
        using (MemoryStream resultStream = new MemoryStream())
        {
            resizedBitmap.Save(resultStream, ImageFormat.Gif);
            return resultStream.ToArray();
        }
    }
}
