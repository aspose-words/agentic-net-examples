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
        // -------------------------------------------------
        // Create a deterministic sample PNG image.
        // -------------------------------------------------
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Use Aspose.Drawing.Bitmap and Aspose.Drawing.Graphics as required.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }

            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // Insert the sample PNG image twice into a Word document.
        // -------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // separate images with a line break
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // Reload the document to simulate a fresh load scenario.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -------------------------------------------------
        // Extract all PNG images, apply contrast enhancement, and save them.
        // -------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // reset before reading

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(imgStream))
                {
                    // Ensure we work with a non‑indexed pixel format.
                    using (Bitmap workingBitmap = new Bitmap(
                        originalBitmap.Width,
                        originalBitmap.Height,
                        PixelFormat.Format32bppArgb))
                    {
                        // Draw the original image onto the new bitmap.
                        using (Graphics g = Graphics.FromImage(workingBitmap))
                        {
                            g.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);
                        }

                        // Apply a simple contrast enhancement.
                        const double contrastFactor = 1.2; // >1 increases contrast
                        for (int y = 0; y < workingBitmap.Height; y++)
                        {
                            for (int x = 0; x < workingBitmap.Width; x++)
                            {
                                Color pixel = workingBitmap.GetPixel(x, y);

                                int r = AdjustContrast(pixel.R, contrastFactor);
                                int gVal = AdjustContrast(pixel.G, contrastFactor);
                                int b = AdjustContrast(pixel.B, contrastFactor);

                                Color newPixel = Color.FromArgb(pixel.A, r, gVal, b);
                                workingBitmap.SetPixel(x, y, newPixel);
                            }
                        }

                        // Save the enhanced image to disk.
                        string outPath = $"extracted-contrast-{extractedCount + 1}.png";
                        workingBitmap.Save(outPath);
                        if (!File.Exists(outPath))
                            throw new InvalidOperationException($"Failed to save enhanced image to '{outPath}'.");
                        extractedCount++;
                    }
                }
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }

    // Helper method to adjust a single channel based on contrast factor.
    private static int AdjustContrast(int value, double factor)
    {
        double newVal = ((value - 128) * factor) + 128;
        return Math.Max(0, Math.Min(255, (int)Math.Round(newVal)));
    }
}
