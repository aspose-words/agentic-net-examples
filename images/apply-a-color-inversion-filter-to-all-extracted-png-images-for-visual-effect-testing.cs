using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 100;
        const int imgHeight = 100;

        // Create a blue square with a red diagonal line.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Blue);
            using (Pen pen = new Pen(Color.Red, 3))
            {
                g.DrawLine(pen, 0, 0, imgWidth - 1, imgHeight - 1);
            }

            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample PNG several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times to have multiple PNG shapes.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // separate images with a line break.
        }

        // Save the document (required by the lifecycle rule).
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract all PNG images, apply a color‑inversion filter, and save them.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;
        int outputIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Obtain raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // Ensure we work with a non‑indexed pixel format.
                    using (Bitmap bitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height, PixelFormat.Format32bppArgb))
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);

                        // Invert colors pixel by pixel.
                        for (int y = 0; y < bitmap.Height; y++)
                        {
                            for (int x = 0; x < bitmap.Width; x++)
                            {
                                Color original = bitmap.GetPixel(x, y);
                                Color inverted = Color.FromArgb(
                                    255 - original.R,
                                    255 - original.G,
                                    255 - original.B);
                                bitmap.SetPixel(x, y, inverted);
                            }
                        }

                        // Save the inverted image.
                        string outFile = Path.Combine(artifactsDir, $"inverted_{outputIndex}.png");
                        bitmap.Save(outFile);
                        outputIndex++;
                    }
                }
            }

            pngCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one PNG was processed.
        // -----------------------------------------------------------------
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        // Optional: verify that at least one output file exists.
        if (!Directory.GetFiles(artifactsDir, "inverted_*.png").Any())
            throw new InvalidOperationException("Inverted images were not saved correctly.");
    }
}
