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
        using (Bitmap bitmap = new Bitmap(300, 300))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple ellipse for visual content.
                g.DrawEllipse(Pens.Black, 50, 50, 200, 200);
            }
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // Step 2: Create a Word document and insert the GIF.
        const string docPath = "DocumentWithGif.docx";
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

                    // Load the GIF using Aspose.Drawing.
                    using (Image originalImage = Image.FromStream(gifStream))
                    {
                        // Create a new bitmap with the target size.
                        using (Bitmap resizedBitmap = new Bitmap(200, 200))
                        {
                            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                            {
                                // Draw the original image scaled to 200x200.
                                graphics.DrawImage(originalImage, 0, 0, 200, 200);
                            }

                            // Save the resized image as a static PNG.
                            string pngPath = $"extracted_{extractedCount}_resized.png";
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
            throw new Exception("No GIF images were found and processed in the document.");
    }
}
