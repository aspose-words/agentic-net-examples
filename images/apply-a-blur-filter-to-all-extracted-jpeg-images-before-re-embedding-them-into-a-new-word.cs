using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging; // For ImageFormat

public class Program
{
    // Simple box blur implementation using Aspose.Drawing types
    private static Bitmap ApplyBoxBlur(Bitmap source, int radius)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap blurred = new Bitmap(width, height);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, count = 0;

                for (int ky = -radius; ky <= radius; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0 || ny >= height) continue;

                    for (int kx = -radius; kx <= radius; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0 || nx >= width) continue;

                        Color pixel = source.GetPixel(nx, ny);
                        rSum += pixel.R;
                        gSum += pixel.G;
                        bSum += pixel.B;
                        count++;
                    }
                }

                int r = rSum / count;
                int g = gSum / count;
                int b = bSum / count;
                blurred.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return blurred;
    }

    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 20, 20, 160, 160);
                }
            }
            // Save as JPEG using Aspose.Drawing.Imaging.ImageFormat
            bmp.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // 2. Create a source Word document containing the JPEG image
        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocPath);

        // 3. Load the source document and process JPEG images
        Document doc = new Document(sourceDocPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Extract image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load image into Bitmap
                using (Bitmap originalBmp = new Bitmap(originalStream))
                {
                    // Apply blur (radius 3)
                    using (Bitmap blurredBmp = ApplyBoxBlur(originalBmp, 3))
                    {
                        // Save blurred image to a new stream (JPEG)
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBmp.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Replace image in the shape
                            shape.ImageData.SetImage(blurredStream);
                        }
                    }
                }
            }
        }

        // 4. Save the modified document
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath);

        // Validation
        if (!File.Exists(outputDocPath))
            throw new Exception("The output document was not created.");

        Console.WriteLine("Processing completed successfully.");
    }
}
