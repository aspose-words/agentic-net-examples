using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Text;

public class Program
{
    public static void Main()
    {
        // ---------- Step 1: Create a sample PNG image ----------
        const string inputImagePath = "input.png";

        // Use Aspose.Drawing.Bitmap explicitly
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(500, 500);
        try
        {
            // Create graphics from the bitmap
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                // Fill background with white
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a blue rectangle
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, 400, 400);
                }
            }
            finally
            {
                g.Dispose();
            }

            // Save the bitmap as PNG
            bitmap.Save(inputImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // ---------- Step 2: Insert the image into a Word document ----------
        const string docPath = "document.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // ---------- Step 3: Extract images from the document ----------
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage) continue;

            // Save the original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading

                // ---------- Step 4: Load the image, resize to 300x300 ----------
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    const int targetSize = 300;
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(targetSize, targetSize))
                    {
                        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                        {
                            g.Clear(Aspose.Drawing.Color.Transparent);
                            // Draw the original image scaled to the target size
                            g.DrawImage(originalBitmap, 0, 0, targetSize, targetSize);

                            // ---------- Step 5: Add watermark text overlay ----------
                            string watermarkText = "Sample Watermark";
                            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                            {
                                // Measure text size
                                SizeF textSize = g.MeasureString(watermarkText, font);
                                // Position text at bottom-right corner with a small margin
                                float x = targetSize - textSize.Width - 10;
                                float y = targetSize - textSize.Height - 10;
                                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Red)))
                                {
                                    g.DrawString(watermarkText, font, brush, x, y);
                                }
                            }
                        }

                        // ---------- Step 6: Save the watermarked image ----------
                        string outputImagePath = $"output_{imageIndex}.png";
                        resizedBitmap.Save(outputImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);

                        // Validate that the file was created
                        if (!File.Exists(outputImagePath))
                            throw new InvalidOperationException($"Failed to create output image: {outputImagePath}");
                    }
                }
            }

            imageIndex++;
        }

        // If no images were processed, throw an exception
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
