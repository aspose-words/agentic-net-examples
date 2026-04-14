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
        // -----------------------------------------------------------------
        // 1. Create a deterministic PNG image (red square) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill the bitmap with red color.
            g.Clear(Aspose.Drawing.Color.Red);
            // Save the bitmap to a local file.
            bitmap.Save(inputImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Insert the image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputImagePath);
        shape.WrapType = WrapType.Inline;
        doc.Save("document.docx");

        // -----------------------------------------------------------------
        // 3. Reload the document and extract PNG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("document.docx");
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape imgShape in shapeNodes.OfType<Shape>())
        {
            if (!imgShape.HasImage) continue;
            if (imgShape.ImageData.ImageType != ImageType.Png) continue;

            // Save the original image to a memory stream.
            using (MemoryStream srcStream = new MemoryStream())
            {
                imgShape.ImageData.Save(srcStream);
                srcStream.Position = 0; // Reset stream position before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap srcBitmap = new Bitmap(srcStream))
                {
                    // -----------------------------------------------------------------
                    // 4. Prepare a hue‑rotation (180°) color matrix.
                    // -----------------------------------------------------------------
                    float lumR = 0.213f, lumG = 0.715f, lumB = 0.072f;
                    // For 180° rotation: cos = -1, sin = 0.
                    float[][] matrixValues = new float[][]
                    {
                        new float[] { 2 * lumR - 1, 2 * lumG,       2 * lumB,       0, 0 },
                        new float[] { 2 * lumR,     2 * lumG - 1,   2 * lumB,       0, 0 },
                        new float[] { 2 * lumR,     2 * lumG,       2 * lumB - 1,   0, 0 },
                        new float[] { 0,            0,              0,              1, 0 },
                        new float[] { 0,            0,              0,              0, 1 }
                    };
                    ColorMatrix hueMatrix = new ColorMatrix(matrixValues);

                    // -----------------------------------------------------------------
                    // 5. Apply the matrix while drawing onto a new bitmap.
                    // -----------------------------------------------------------------
                    using (Bitmap destBitmap = new Bitmap(srcBitmap.Width, srcBitmap.Height))
                    using (Graphics g = Graphics.FromImage(destBitmap))
                    using (ImageAttributes attr = new ImageAttributes())
                    {
                        attr.SetColorMatrix(hueMatrix);
                        g.DrawImage(
                            srcBitmap,
                            new Rectangle(0, 0, destBitmap.Width, destBitmap.Height),
                            0, 0, srcBitmap.Width, srcBitmap.Height,
                            GraphicsUnit.Pixel,
                            attr);

                        // Save the transformed image.
                        string outPath = $"extracted_{extractedCount}_rotated.png";
                        destBitmap.Save(outPath);
                        extractedCount++;
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 6. Validation – ensure at least one image was processed.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Optional clean‑up (commented out).
        // File.Delete(inputImagePath);
        // File.Delete("document.docx");
    }
}
