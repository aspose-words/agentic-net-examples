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
        // Prepare deterministic file names
        const string sampleImagePath = "input.png";
        const string docPath = "sample.docx";
        const string outputDocPath = "output.docx";
        const string sharpenedImagePath = "sharpened.png";

        // 1. Create a sample PNG image
        CreateSamplePng(sampleImagePath);

        // 2. Create a Word document and insert the PNG image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // 3. Load the document for processing
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Ensure the image is a PNG (by checking the image bytes header)
            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (imageBytes == null || imageBytes.Length < 8)
                continue;

            // PNG signature: 89 50 4E 47 0D 0A 1A 0A
            if (imageBytes[0] != 0x89 || imageBytes[1] != 0x50 || imageBytes[2] != 0x4E ||
                imageBytes[3] != 0x47)
                continue; // Not a PNG

            // 4. Save the extracted PNG to a temporary file
            const string extractedPath = "extracted.png";
            shape.ImageData.Save(extractedPath);

            // 5. Load the PNG, apply sharpening filter, and save
            using (Bitmap bitmap = new Bitmap(extractedPath))
            {
                using (Bitmap sharpened = ApplySharpenFilter(bitmap))
                {
                    sharpened.Save(sharpenedImagePath, ImageFormat.Png);
                }
            }

            // 6. Replace the shape's image with the sharpened version
            shape.ImageData.SetImage(sharpenedImagePath);
            processedCount++;
        }

        // Validate that at least one image was processed
        if (processedCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed.");

        // 7. Save the modified document
        loadedDoc.Save(outputDocPath);

        // Validate output file existence
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        Console.WriteLine($"Processing complete. {processedCount} image(s) sharpened and saved to '{outputDocPath}'.");
    }

    private static void CreateSamplePng(string path)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
                }
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 70, 70, 60, 60);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static Bitmap ApplySharpenFilter(Bitmap source)
    {
        int w = source.Width;
        int h = source.Height;
        Bitmap result = new Bitmap(w, h);

        // Sharpen kernel
        int[,] kernel = {
            {  0, -1,  0 },
            { -1,  5, -1 },
            {  0, -1,  0 }
        };
        int kernelSize = 3;
        int offset = kernelSize / 2;

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                int r = 0, g = 0, b = 0;

                for (int ky = -offset; ky <= offset; ky++)
                {
                    int py = y + ky;
                    if (py < 0 || py >= h) continue;

                    for (int kx = -offset; kx <= offset; kx++)
                    {
                        int px = x + kx;
                        if (px < 0 || px >= w) continue;

                        Color pixelColor = source.GetPixel(px, py);
                        int weight = kernel[ky + offset, kx + offset];

                        r += pixelColor.R * weight;
                        g += pixelColor.G * weight;
                        b += pixelColor.B * weight;
                    }
                }

                // Clamp values to byte range
                r = Math.Min(Math.Max(r, 0), 255);
                g = Math.Min(Math.Max(g, 0), 255);
                b = Math.Min(Math.Max(b, 0), 255);

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}
