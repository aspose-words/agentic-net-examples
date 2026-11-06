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
        // Create a sample GIF image (100x100) with deterministic content.
        const string gifPath = "sample.gif";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                g.FillRectangle(new SolidBrush(Color.DarkBlue), 20, 20, 60, 60);
            }
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // Ensure the GIF file exists.
        if (!File.Exists(gifPath))
            throw new Exception("Failed to create sample GIF.");

        // Create a Word document and insert the GIF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        const string docPath = "doc_with_gif.docx";
        doc.Save(docPath);

        // Load the document (optional, as we already have it in memory).
        Document loadedDoc = new Document(docPath);

        // Iterate through all Shape nodes to find GIF images.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Extract the GIF image to a memory stream.
                using (MemoryStream gifStream = new MemoryStream())
                {
                    shape.ImageData.Save(gifStream);
                    gifStream.Position = 0;

                    // Load the GIF into a Bitmap.
                    using (Bitmap originalBitmap = new Bitmap(gifStream))
                    {
                        // Create a new 200x200 bitmap for the resized image.
                        using (Bitmap resizedBitmap = new Bitmap(200, 200))
                        {
                            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                            {
                                graphics.Clear(Color.White);
                                // Draw the original image scaled to 200x200.
                                graphics.DrawImage(originalBitmap, 0, 0, 200, 200);
                            }

                            // Save the resized image as a static PNG.
                            const string pngPath = "extracted_resized.png";
                            resizedBitmap.Save(pngPath, ImageFormat.Png);

                            // Validate that the PNG file was created.
                            if (!File.Exists(pngPath))
                                throw new Exception("Resized PNG image was not created.");
                        }
                    }
                }
            }
        }
    }
}
