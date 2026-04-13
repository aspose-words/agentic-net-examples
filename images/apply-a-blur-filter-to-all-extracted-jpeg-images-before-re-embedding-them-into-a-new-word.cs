using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string inputImagePath = Path.Combine(artifactsDir, "input.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            // Save as JPEG.
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a source Word document that contains the sample image.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.InsertImage(inputImagePath);
        srcBuilder.InsertParagraph();
        srcBuilder.InsertImage(inputImagePath);
        string srcDocPath = Path.Combine(artifactsDir, "Source.docx");
        srcDoc.Save(srcDocPath);

        // -----------------------------------------------------------------
        // 3. Load the source document, extract JPEG images, blur them,
        //    and embed the blurred images into a new document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(srcDocPath);
        Document newDoc = new Document();
        DocumentBuilder newBuilder = new DocumentBuilder(newDoc);

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image into a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply a simple box blur.
                    using (Bitmap blurredBitmap = ApplyBoxBlur(originalBitmap))
                    {
                        // Save the blurred bitmap back to a stream.
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Insert the blurred image into the new document.
                            newBuilder.InsertImage(blurredStream);
                        }
                    }
                }
            }
        }

        // Save the new document.
        string newDocPath = Path.Combine(artifactsDir, "BlurredImages.docx");
        newDoc.Save(newDocPath);

        // Validate that the output file was created.
        if (!File.Exists(newDocPath))
            throw new Exception("The output document was not created.");
    }

    // -----------------------------------------------------------------
    // Simple box blur implementation for a bitmap.
    // -----------------------------------------------------------------
    private static Bitmap ApplyBoxBlur(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap result = new Bitmap(width, height);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int sumR = 0, sumG = 0, sumB = 0, count = 0;

                for (int offsetY = -1; offsetY <= 1; offsetY++)
                {
                    int ny = y + offsetY;
                    if (ny < 0 || ny >= height)
                        continue;

                    for (int offsetX = -1; offsetX <= 1; offsetX++)
                    {
                        int nx = x + offsetX;
                        if (nx < 0 || nx >= width)
                            continue;

                        Color pixel = source.GetPixel(nx, ny);
                        sumR += pixel.R;
                        sumG += pixel.G;
                        sumB += pixel.B;
                        count++;
                    }
                }

                Color blurred = Color.FromArgb(sumR / count, sumG / count, sumB / count);
                result.SetPixel(x, y, blurred);
            }
        }

        return result;
    }
}
