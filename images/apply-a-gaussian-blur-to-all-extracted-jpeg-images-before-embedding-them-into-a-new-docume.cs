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
        // Prepare deterministic file names.
        string workDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(workDir, "sample.jpg");
        string sourceDocPath = Path.Combine(workDir, "source.docx");
        string resultDocPath = Path.Combine(workDir, "result.docx");

        // -------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 150;
        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple red rectangle.
            using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save as JPEG.
            bmp.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -------------------------------------------------
        // 2. Create a source document and insert the JPEG several times.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document with original images:");
        srcBuilder.InsertImage(sampleImagePath);
        srcBuilder.InsertParagraph();
        srcBuilder.InsertImage(sampleImagePath);
        srcBuilder.InsertParagraph();
        srcBuilder.InsertImage(sampleImagePath);
        sourceDoc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the source document and extract JPEG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        // Prepare a new document where blurred images will be inserted.
        Document resultDoc = new Document();
        DocumentBuilder resBuilder = new DocumentBuilder(resultDoc);
        resBuilder.Writeln("Result document with Gaussian‑blurred images:");

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Obtain the original image bytes.
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
            {
                // Apply Gaussian blur.
                using (Aspose.Drawing.Bitmap blurredBitmap = ApplyGaussianBlur(originalBitmap))
                {
                    // Save blurred bitmap to a memory stream (JPEG format).
                    using (MemoryStream blurredStream = new MemoryStream())
                    {
                        blurredBitmap.Save(blurredStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        blurredStream.Position = 0;

                        // Insert the blurred image into the result document.
                        resBuilder.InsertImage(blurredStream);
                        resBuilder.InsertParagraph();
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Save the result document.
        // -------------------------------------------------
        resultDoc.Save(resultDocPath);

        // Validate that the output file was created.
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("Result document was not created.");
    }

    // -------------------------------------------------
    // Helper: Apply a simple Gaussian blur to a bitmap.
    // -------------------------------------------------
    private static Aspose.Drawing.Bitmap ApplyGaussianBlur(Aspose.Drawing.Bitmap source)
    {
        const int radius = 2;               // Kernel radius (2 => 5x5 kernel)
        const double sigma = 1.0;           // Standard deviation

        int size = radius * 2 + 1;
        double[,] kernel = new double[size, size];
        double kernelSum = 0.0;

        // Build Gaussian kernel.
        for (int y = -radius; y <= radius; y++)
        {
            for (int x = -radius; x <= radius; x++)
            {
                double exponent = -(x * x + y * y) / (2 * sigma * sigma);
                double value = Math.Exp(exponent);
                kernel[y + radius, x + radius] = value;
                kernelSum += value;
            }
        }

        // Normalize kernel.
        for (int y = 0; y < size; y++)
            for (int x = 0; x < size; x++)
                kernel[y, x] /= kernelSum;

        int width = source.Width;
        int height = source.Height;
        Aspose.Drawing.Bitmap blurred = new Aspose.Drawing.Bitmap(width, height);

        // Convolution.
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                double r = 0, g = 0, b = 0;

                for (int ky = -radius; ky <= radius; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0) ny = 0;
                    if (ny >= height) ny = height - 1;

                    for (int kx = -radius; kx <= radius; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0) nx = 0;
                        if (nx >= width) nx = width - 1;

                        Aspose.Drawing.Color pixel = source.GetPixel(nx, ny);
                        double weight = kernel[ky + radius, kx + radius];

                        r += pixel.R * weight;
                        g += pixel.G * weight;
                        b += pixel.B * weight;
                    }
                }

                int ir = Math.Min(255, Math.Max(0, (int)Math.Round(r)));
                int ig = Math.Min(255, Math.Max(0, (int)Math.Round(g)));
                int ib = Math.Min(255, Math.Max(0, (int)Math.Round(b)));

                blurred.SetPixel(x, y, Aspose.Drawing.Color.FromArgb(ir, ig, ib));
            }
        }

        return blurred;
    }
}
