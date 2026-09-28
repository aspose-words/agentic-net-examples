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
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string outputImagesDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(outputImagesDir);

        // Create a deterministic sample PNG image
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents that contain the PNG image
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), sampleImagePath);

        // Process each document: extract PNG images, apply sepia, save result
        int totalProcessed = 0;
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                // Only process PNG images
                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Save the original image to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading

                    // Load bitmap from stream (may be indexed)
                    using (Bitmap originalBitmap = new Bitmap(imgStream))
                    {
                        // Convert to a non‑indexed bitmap to allow pixel manipulation
                        using (Bitmap workingBitmap = ConvertToNonIndexed(originalBitmap))
                        {
                            ApplySepiaTone(workingBitmap);

                            // Prepare output file name
                            string docName = Path.GetFileNameWithoutExtension(docPath);
                            string outFileName = $"{docName}_image{imageIndex}_sepia.png";
                            string outPath = Path.Combine(outputImagesDir, outFileName);

                            // Save the sepia image
                            workingBitmap.Save(outPath, ImageFormat.Png);
                            totalProcessed++;
                        }
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image was processed
        if (totalProcessed == 0)
            throw new InvalidOperationException("No PNG images were found and processed.");

        Console.WriteLine($"Processing complete. {totalProcessed} image(s) saved to '{outputImagesDir}'.");
    }

    // Creates a simple deterministic PNG image file
    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a Word document with the specified PNG image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Converts an indexed bitmap to a 24‑bpp RGB bitmap
    private static Bitmap ConvertToNonIndexed(Bitmap source)
    {
        Bitmap dest = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
        using (Graphics g = Graphics.FromImage(dest))
        {
            g.DrawImage(source, 0, 0, source.Width, source.Height);
        }
        return dest;
    }

    // Applies a sepia tone effect to the provided bitmap
    private static void ApplySepiaTone(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                int r = original.R;
                int g = original.G;
                int b = original.B;

                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                tr = Math.Min(255, tr);
                tg = Math.Min(255, tg);
                tb = Math.Min(255, tb);

                Color sepia = Color.FromArgb(tr, tg, tb);
                bitmap.SetPixel(x, y, sepia);
            }
        }
    }
}
