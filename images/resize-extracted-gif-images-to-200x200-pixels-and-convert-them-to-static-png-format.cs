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
        // Step 1: Create a sample GIF image.
        const string gifPath = "sample.gif";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // Step 2: Create a Word document and insert the GIF.
        const string docPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // Step 3: Load the document and extract GIF images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Save the original GIF to a memory stream.
                using (MemoryStream gifStream = new MemoryStream())
                {
                    shape.ImageData.Save(gifStream);
                    gifStream.Position = 0;

                    // Load the GIF into a bitmap.
                    using (Bitmap originalBitmap = new Bitmap(gifStream))
                    {
                        // Create a new bitmap with the target size.
                        using (Bitmap resizedBitmap = new Bitmap(200, 200))
                        {
                            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                            {
                                graphics.Clear(Color.White);
                                graphics.DrawImage(originalBitmap, 0, 0, 200, 200);
                            }

                            // Save the resized image as PNG.
                            string pngPath = $"extracted_{extractedCount}.png";
                            resizedBitmap.Save(pngPath, ImageFormat.Png);

                            // Validate that the PNG was created.
                            if (!File.Exists(pngPath))
                                throw new Exception($"Failed to create PNG file: {pngPath}");

                            extractedCount++;
                        }
                    }
                }
            }
        }

        // Validate that at least one image was processed.
        if (extractedCount == 0)
            throw new Exception("No GIF images were found and extracted.");

        // Cleanup sample files (optional).
        // File.Delete(gifPath);
        // File.Delete(docPath);
    }
}
