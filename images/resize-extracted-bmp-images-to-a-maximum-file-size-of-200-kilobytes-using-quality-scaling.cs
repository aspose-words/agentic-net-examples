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
        // Create a sample BMP image.
        const string inputBmpPath = "sample.bmp";
        const int originalWidth = 800;
        const int originalHeight = 600;
        Bitmap bmp = new Bitmap(originalWidth, originalHeight);
        Graphics g = Graphics.FromImage(bmp);
        g.Clear(Color.White);
        // Draw a simple rectangle to have some content.
        g.FillRectangle(new SolidBrush(Color.Blue), 100, 100, 600, 400);
        g.Dispose();
        bmp.Save(inputBmpPath, ImageFormat.Bmp);
        bmp.Dispose();

        // Create a Word document and insert the BMP image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputBmpPath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates extraction).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Extract image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the extracted BMP into Aspose.Drawing.Bitmap.
                Bitmap extractedBmp = new Bitmap(imageStream);
                // Resize loop: reduce dimensions until file size <= 200 KB.
                const long maxFileSize = 200 * 1024; // 200 KB
                while (true)
                {
                    using (MemoryStream tempStream = new MemoryStream())
                    {
                        extractedBmp.Save(tempStream, ImageFormat.Bmp);
                        if (tempStream.Length <= maxFileSize) break;
                    }

                    // Reduce dimensions by 90%.
                    int newWidth = (int)(extractedBmp.Width * 0.9);
                    int newHeight = (int)(extractedBmp.Height * 0.9);
                    if (newWidth < 1 || newHeight < 1) break; // Prevent zero size.

                    Bitmap resizedBmp = new Bitmap(newWidth, newHeight);
                    Graphics graphics = Graphics.FromImage(resizedBmp);
                    graphics.DrawImage(extractedBmp, 0, 0, newWidth, newHeight);
                    graphics.Dispose();
                    extractedBmp.Dispose();
                    extractedBmp = resizedBmp;
                }

                // Save the resized BMP.
                string outputPath = $"resized_image_{imageIndex}.bmp";
                extractedBmp.Save(outputPath, ImageFormat.Bmp);
                extractedBmp.Dispose();

                // Validate output.
                if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                {
                    throw new InvalidOperationException($"Failed to create resized image: {outputPath}");
                }

                imageIndex++;
            }
        }

        // Final validation: at least one image should have been processed.
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were found and processed in the document.");
        }
    }
}
