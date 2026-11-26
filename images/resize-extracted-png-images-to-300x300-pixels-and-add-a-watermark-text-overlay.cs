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
        // Create a deterministic sample PNG image (200x200) with a blue ellipse.
        const string inputImagePath = "input.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
        {
            g.FillEllipse(brush, new Aspose.Drawing.Rectangle(20, 20, 160, 160));
        }
        bitmap.Save(inputImagePath);
        g.Dispose();
        bitmap.Dispose();

        // Verify the sample image was created.
        if (!File.Exists(inputImagePath))
            throw new Exception("Failed to create the input image.");

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Verify the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("Failed to create the sample document.");

        // Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image to a memory stream.
            using (MemoryStream extractedStream = new MemoryStream())
            {
                shape.ImageData.Save(extractedStream);
                extractedStream.Position = 0;

                // Load the extracted image into a bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(extractedStream))
                {
                    const int targetSize = 300;
                    // Create a new bitmap for the resized image.
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(targetSize, targetSize))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Aspose.Drawing.Color.White);
                            // Draw the original image scaled to the target size.
                            graphics.DrawImage(originalBitmap, new Aspose.Drawing.Rectangle(0, 0, targetSize, targetSize));

                            // Add watermark text using Aspose.Drawing.Font explicitly.
                            using (Aspose.Drawing.Font watermarkFont = new Aspose.Drawing.Font("Arial", 20))
                            using (Aspose.Drawing.SolidBrush watermarkBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                            {
                                string watermarkText = "Watermark";
                                // Position the watermark near the bottom-left corner.
                                Aspose.Drawing.PointF position = new Aspose.Drawing.PointF(10, targetSize - 30);
                                graphics.DrawString(watermarkText, watermarkFont, watermarkBrush, position);
                            }
                        }

                        // Save the watermarked image.
                        string outputImagePath = $"watermarked-{imageIndex}.png";
                        resizedBitmap.Save(outputImagePath);
                        if (!File.Exists(outputImagePath))
                            throw new Exception($"Failed to save watermarked image: {outputImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one watermarked image was produced.
        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional cleanup (commented out).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
