using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.jpg";
        const string originalDocPath = "original.docx";
        const string outputDocPath = "blurred_output.docx";

        // 1. Create a deterministic sample JPEG image.
        CreateSampleJpeg(sampleImagePath);

        // 2. Create a Word document and insert the sample JPEG image.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.InsertImage(sampleImagePath);
        originalDoc.Save(originalDocPath);

        // 3. Load the document, extract JPEG images, apply blur, and collect blurred image paths.
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        List<string> blurredImagePaths = new List<string>();

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            // Check if the image is JPEG.
            ImageType imageType = shape.ImageData.ImageType;
            if (imageType != ImageType.Jpeg) continue;

            // Extract image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load bitmap from stream.
                using (Bitmap bitmap = new Bitmap(imageStream))
                {
                    // Apply a simple box blur.
                    using (Bitmap blurredBitmap = ApplyBoxBlur(bitmap))
                    {
                        // Save blurred image to a deterministic file.
                        string blurredPath = $"blurred_{imageIndex}.jpg";
                        blurredBitmap.Save(blurredPath, ImageFormat.Jpeg);
                        blurredImagePaths.Add(blurredPath);
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one blurred image was created.
        if (blurredImagePaths.Count == 0)
            throw new InvalidOperationException("No JPEG images were found to blur.");

        // 4. Create a new document and embed the blurred images.
        Document outputDoc = new Document();
        DocumentBuilder outBuilder = new DocumentBuilder(outputDoc);

        foreach (string blurredPath in blurredImagePaths)
        {
            outBuilder.InsertImage(blurredPath);
            outBuilder.Writeln(); // Add a paragraph break between images.
        }

        outputDoc.Save(outputDocPath);

        // Validate that the output document exists.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(originalDocPath);
        // foreach (var path in blurredImagePaths) File.Delete(path);
    }

    // Creates a simple deterministic JPEG image.
    private static void CreateSampleJpeg(string filePath)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, 160, 60);
                }
            }
            // Save explicitly as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Applies a simple 3x3 box blur to the provided bitmap and returns a new blurred bitmap.
    private static Bitmap ApplyBoxBlur(Bitmap source)
    {
        int w = source.Width;
        int h = source.Height;
        Bitmap blurred = new Bitmap(w, h);

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                int rSum = 0, gSum = 0, bSum = 0, count = 0;

                // Iterate over the 3x3 kernel.
                for (int ky = -1; ky <= 1; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0 || ny >= h) continue;

                    for (int kx = -1; kx <= 1; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0 || nx >= w) continue;

                        Color neighbor = source.GetPixel(nx, ny);
                        rSum += neighbor.R;
                        gSum += neighbor.G;
                        bSum += neighbor.B;
                        count++;
                    }
                }

                // Average the colors.
                Color avg = Color.FromArgb(rSum / count, gSum / count, bSum / count);
                blurred.SetPixel(x, y, avg);
            }
        }

        return blurred;
    }
}
