using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        string sampleImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen bluePen = new Pen(Color.Blue, 2))
                {
                    graphics.DrawEllipse(bluePen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = "DocumentWithImages.docx";
        doc.Save(docPath);

        // Load the document and extract all PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Save the original image to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0;

                    // Load the image into Aspose.Drawing.Bitmap.
                    using (Bitmap bmp = new Bitmap(imageStream))
                    {
                        // Apply a 5‑pixel red border.
                        using (Graphics gfx = Graphics.FromImage(bmp))
                        {
                            using (Pen redPen = new Pen(Color.Red, 5))
                            {
                                Rectangle rect = new Rectangle(0, 0, bmp.Width - 1, bmp.Height - 1);
                                gfx.DrawRectangle(redPen, rect);
                            }
                        }

                        // Save the modified image.
                        string outputPath = $"extracted_{extractedCount}.png";
                        bmp.Save(outputPath, ImageFormat.Png);
                        if (!File.Exists(outputPath))
                            throw new Exception($"Failed to save modified image: {outputPath}");

                        extractedCount++;
                    }
                }
            }
        }

        if (extractedCount == 0)
            throw new Exception("No PNG images were extracted from the document.");
    }
}
