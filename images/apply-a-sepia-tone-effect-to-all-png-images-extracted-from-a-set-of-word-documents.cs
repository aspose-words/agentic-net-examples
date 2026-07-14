using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class SepiaImageProcessor
{
    public static void Main()
    {
        // Prepare directories.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string inputDocsDir = Path.Combine(artifactsDir, "InputDocs");
        Directory.CreateDirectory(inputDocsDir);
        string outputImagesDir = Path.Combine(artifactsDir, "SepiaImages");
        Directory.CreateDirectory(outputImagesDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -----------------------------------------------------------------
        string samplePngPath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Fill background.
                g.Clear(Aspose.Drawing.Color.LightGray);
                // Draw a simple red ellipse.
                using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillEllipse(brush, 25, 25, 150, 150);
                }
            }
            bmp.Save(samplePngPath);
        }

        // -----------------------------------------------------------------
        // 2. Create a few Word documents that contain the PNG image.
        // -----------------------------------------------------------------
        int docCount = 2;
        for (int i = 0; i < docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i + 1}");
            // Insert the sample PNG.
            builder.InsertImage(samplePngPath);
            string docPath = Path.Combine(inputDocsDir, $"Doc{i + 1}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Extract PNG images, apply sepia tone, and save them.
        // -----------------------------------------------------------------
        int processedImageCount = 0;
        string[] docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Process only PNG images.
                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Obtain image bytes.
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load into Aspose.Drawing.Bitmap and ensure a non‑indexed pixel format.
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    using (Bitmap original = new Bitmap(ms))
                    {
                        // Create a 32‑bpp ARGB bitmap to allow SetPixel.
                        using (Bitmap bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb))
                        {
                            using (Graphics g = Graphics.FromImage(bitmap))
                            {
                                g.DrawImage(original, 0, 0, original.Width, original.Height);
                            }

                            ApplySepia(bitmap);

                            // Build output file name.
                            string outFileName = $"sepia_{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.png";
                            string outPath = Path.Combine(outputImagesDir, outFileName);
                            bitmap.Save(outPath);
                            processedImageCount++;
                            imageIndex++;
                        }
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Validation.
        // -----------------------------------------------------------------
        if (processedImageCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed.");

        Console.WriteLine($"Processed {processedImageCount} PNG image(s). Sepia images saved to: {outputImagesDir}");
    }

    // Applies a sepia tone effect to the provided bitmap.
    private static void ApplySepia(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);

                int tr = (int)(original.R * 0.393 + original.G * 0.769 + original.B * 0.189);
                int tg = (int)(original.R * 0.349 + original.G * 0.686 + original.B * 0.168);
                int tb = (int)(original.R * 0.272 + original.G * 0.534 + original.B * 0.131);

                // Clamp values to byte range.
                byte r = (byte)(tr > 255 ? 255 : tr);
                byte g = (byte)(tg > 255 ? 255 : tg);
                byte b = (byte)(tb > 255 ? 255 : tb);

                Aspose.Drawing.Color sepia = Aspose.Drawing.Color.FromArgb(r, g, b);
                bitmap.SetPixel(x, y, sepia);
            }
        }
    }
}
