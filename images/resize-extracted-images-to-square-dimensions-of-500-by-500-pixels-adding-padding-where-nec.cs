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
        // Create a deterministic sample image (input.png)
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(300, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(Brushes.Blue, 50, 30, 200, 140);
            }
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the sample image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load step)
        Document loadedDoc = new Document(docPath);

        // Extract images, resize to 500x500 with padding, and save
        const int targetSize = 500;
        int imageIndex = 0;
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            // Get original image bytes
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure stream is at start
                using (Bitmap original = new Bitmap(ms))
                {
                    // Calculate scaled dimensions while preserving aspect ratio
                    float ratio = Math.Min((float)targetSize / original.Width, (float)targetSize / original.Height);
                    int newWidth = (int)(original.Width * ratio);
                    int newHeight = (int)(original.Height * ratio);
                    int offsetX = (targetSize - newWidth) / 2;
                    int offsetY = (targetSize - newHeight) / 2;

                    // Create a new square bitmap with white background
                    using (Bitmap resized = new Bitmap(targetSize, targetSize))
                    {
                        using (Graphics g = Graphics.FromImage(resized))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(original, offsetX, offsetY, newWidth, newHeight);
                        }

                        // Save the resized image
                        string outputImagePath = $"extracted_resized_{imageIndex}.png";
                        resized.Save(outputImagePath, ImageFormat.Png);
                        if (!File.Exists(outputImagePath))
                            throw new InvalidOperationException($"Failed to create output image: {outputImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Clean up temporary files (optional)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
