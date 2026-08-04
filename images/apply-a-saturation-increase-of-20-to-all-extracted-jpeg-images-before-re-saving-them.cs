using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageSaturationExample
{
    // Increases the saturation of a bitmap by the given factor (e.g., 1.2 for +20%).
    private static Bitmap IncreaseSaturation(Bitmap source, float factor)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);
        using (Graphics graphics = Graphics.FromImage(result))
        {
            // Build a saturation color matrix.
            float sat = factor;
            float invSat = 1f - sat;
            float rWeight = 0.2126f * invSat;
            float gWeight = 0.7152f * invSat;
            float bWeight = 0.0722f * invSat;

            float[][] matrixElements = new float[][]
            {
                new float[] { rWeight + sat, rWeight,          rWeight,          0, 0 },
                new float[] { gWeight,        gWeight + sat,    gWeight,          0, 0 },
                new float[] { bWeight,        bWeight,          bWeight + sat,    0, 0 },
                new float[] { 0,              0,                0,                1, 0 },
                new float[] { 0,              0,                0,                0, 1 }
            };

            ColorMatrix colorMatrix = new ColorMatrix(matrixElements);
            ImageAttributes attributes = new ImageAttributes();
            attributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            // Draw the original image onto the new bitmap using the color matrix.
            graphics.DrawImage(
                source,
                new Rectangle(0, 0, width, height),
                0,
                0,
                width,
                height,
                GraphicsUnit.Pixel,
                attributes);
        }
        return result;
    }

    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image that will be inserted into the document.
        // -----------------------------------------------------------------
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                g.FillEllipse(Aspose.Drawing.Brushes.Crimson, 25, 25, 150, 150);
            }
            bmp.Save(sampleJpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with JPEG images:");
        builder.InsertImage(sampleJpegPath);
        builder.InsertParagraph();
        builder.InsertImage(sampleJpegPath);
        builder.InsertParagraph();
        builder.InsertImage(sampleJpegPath);
        string docPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, find JPEG images, increase their saturation by 20%,
        //    and save the modified images back to the document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load into a bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap originalBmp = new Bitmap(ms))
            {
                // Increase saturation by 20% (factor 1.2).
                using (Bitmap saturatedBmp = IncreaseSaturation(originalBmp, 1.2f))
                {
                    // Save the modified bitmap to a new memory stream.
                    using (MemoryStream outMs = new MemoryStream())
                    {
                        saturatedBmp.Save(outMs, ImageFormat.Jpeg);
                        outMs.Position = 0;

                        // Replace the image in the shape.
                        shape.ImageData.SetImage(outMs);
                    }

                    // Also save the saturated image as a separate file for verification.
                    string extractedPath = Path.Combine(artifactsDir, $"Extracted_{jpegCount}.jpg");
                    saturatedBmp.Save(extractedPath, ImageFormat.Jpeg);
                }
            }

            jpegCount++;
        }

        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        // -----------------------------------------------------------------
        // 4. Save the document with the updated images.
        // -----------------------------------------------------------------
        string updatedDocPath = Path.Combine(artifactsDir, "Updated.docx");
        loadedDoc.Save(updatedDocPath);

        Console.WriteLine($"Processed {jpegCount} JPEG image(s).");
        Console.WriteLine($"Original document: {docPath}");
        Console.WriteLine($"Updated document : {updatedDocPath}");
        Console.WriteLine($"Extracted images are saved in: {artifactsDir}");
    }
}
