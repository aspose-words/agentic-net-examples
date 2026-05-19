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
        string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath);

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document (demonstrates load/save lifecycle).
        Document loadedDoc = new Document(docPath);

        // Extract all PNG images and apply a hue rotation of 180 degrees.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Save the shape image to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset stream before reading.

                    // Load the image into an Aspose.Drawing.Bitmap.
                    using (Bitmap originalBitmap = new Bitmap(ms))
                    // Apply the hue rotation.
                    using (Bitmap rotatedBitmap = ApplyHueRotation(originalBitmap, 180f))
                    {
                        string outputPath = $"output_{imageIndex}.png";
                        rotatedBitmap.Save(outputPath, ImageFormat.Png);
                        if (!File.Exists(outputPath))
                            throw new Exception($"Failed to save {outputPath}");
                    }
                }

                imageIndex++;
            }
        }

        // Validation: at least one PNG image must have been processed.
        if (imageIndex == 0)
            throw new Exception("No PNG images were found in the document.");
    }

    // Generates a simple PNG image with distinct colors for testing.
    private static void CreateSamplePng(string path)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                graphics.FillRectangle(brush, 50, 50, 100, 100);
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Applies a hue rotation (in degrees) to the supplied bitmap and returns a new bitmap.
    private static Bitmap ApplyHueRotation(Bitmap source, float degrees)
    {
        float angle = degrees * (float)Math.PI / 180f;
        float cos = (float)Math.Cos(angle);
        float sin = (float)Math.Sin(angle);

        const float lumR = 0.213f;
        const float lumG = 0.715f;
        const float lumB = 0.072f;

        // Build the hue rotation matrix.
        float[][] matrix = new float[5][]
        {
            new float[]
            {
                lumR + cos * (1 - lumR) + sin * (-lumR),
                lumG + cos * (-lumG) + sin * (-lumG),
                lumB + cos * (-lumB) + sin * (1 - lumB),
                0, 0
            },
            new float[]
            {
                lumR + cos * (-lumR) + sin * 0.143f,
                lumG + cos * (1 - lumG) + sin * 0.140f,
                lumB + cos * (-lumB) + sin * (-0.283f),
                0, 0
            },
            new float[]
            {
                lumR + cos * (-lumR) + sin * (-(1 - lumR)),
                lumG + cos * (-lumG) + sin * lumG,
                lumB + cos * (1 - lumB) + sin * lumB,
                0, 0
            },
            new float[] { 0, 0, 0, 1, 0 },
            new float[] { 0, 0, 0, 0, 1 }
        };

        ColorMatrix colorMatrix = new ColorMatrix(matrix);
        ImageAttributes attributes = new ImageAttributes();
        attributes.SetColorMatrix(colorMatrix);

        Bitmap result = new Bitmap(source.Width, source.Height);
        using (Graphics g = Graphics.FromImage(result))
        {
            Rectangle rect = new Rectangle(0, 0, source.Width, source.Height);
            g.DrawImage(source, rect, 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes);
        }

        attributes.Dispose();
        return result;
    }
}
