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
        // File names
        const string gifPath = "sample.gif";
        const string docPath = "doc_with_gif.docx";
        const string outputPngPath = "extracted_resized.png";

        // -------------------------------------------------
        // 1. Create a sample 400x400 GIF image using Aspose.Drawing
        // -------------------------------------------------
        const int originalSize = 400;
        using (Bitmap bitmap = new Bitmap(originalSize, originalSize))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with light blue
                g.Clear(Color.LightBlue);
                // Draw a red ellipse
                g.FillEllipse(new SolidBrush(Color.Red), 50, 50, 300, 300);
            }

            // Save as GIF
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the GIF image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract the GIF image
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        bool extracted = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Save image data to a memory stream
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset stream position

                    // Load the GIF into a bitmap
                    using (Bitmap originalBitmap = new Bitmap(ms))
                    {
                        const int targetSize = 200;
                        using (Bitmap resizedBitmap = new Bitmap(targetSize, targetSize))
                        {
                            using (Graphics g = Graphics.FromImage(resizedBitmap))
                            {
                                // Draw the original image scaled to the new size
                                g.DrawImage(originalBitmap, 0, 0, targetSize, targetSize);
                            }

                            // Save the resized bitmap as PNG
                            resizedBitmap.Save(outputPngPath, ImageFormat.Png);
                            extracted = true;
                        }
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Validate that the PNG file was created
        // -------------------------------------------------
        if (!extracted || !File.Exists(outputPngPath))
        {
            throw new Exception("Failed to extract and resize GIF image to PNG.");
        }

        // Optional cleanup (commented out)
        // File.Delete(gifPath);
        // File.Delete(docPath);
    }
}
