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
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            bmp.Save(sampleImagePath);
        }

        // Step 2: Create a Word document and insert the PNG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Reload the document (simulating a separate operation).
        Document loadedDoc = new Document(docPath);

        // Step 4: Extract all PNG images, enhance contrast, and save them.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Verify the image is a PNG using ImageType.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;

                // Load the image into a bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(ms))
                {
                    // Prepare contrast enhancement.
                    const float contrastFactor = 1.2f; // Increase contrast by 20%
                    float t = (1.0f - contrastFactor) / 2.0f;
                    float[][] matrixValues = new float[][]
                    {
                        new float[] { contrastFactor, 0, 0, 0, 0 },
                        new float[] { 0, contrastFactor, 0, 0, 0 },
                        new float[] { 0, 0, contrastFactor, 0, 0 },
                        new float[] { 0, 0, 0, 1, 0 },
                        new float[] { t, t, t, 0, 1 }
                    };
                    ColorMatrix contrastMatrix = new ColorMatrix(matrixValues);
                    ImageAttributes imgAttr = new ImageAttributes();
                    imgAttr.SetColorMatrix(contrastMatrix);

                    // Create a new bitmap to hold the enhanced image.
                    using (Aspose.Drawing.Bitmap enhancedBitmap = new Aspose.Drawing.Bitmap(originalBitmap.Width, originalBitmap.Height))
                    {
                        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(enhancedBitmap))
                        {
                            g.DrawImage(
                                originalBitmap,
                                new Rectangle(0, 0, enhancedBitmap.Width, enhancedBitmap.Height),
                                0,
                                0,
                                originalBitmap.Width,
                                originalBitmap.Height,
                                GraphicsUnit.Pixel,
                                imgAttr);
                        }

                        // Save the enhanced image.
                        string outputPath = $"extracted-{imageIndex}-enhanced.png";
                        enhancedBitmap.Save(outputPath);
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to save enhanced image to '{outputPath}'.");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found and processed in the document.");
    }
}
