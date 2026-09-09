using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File names
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // 1. Create a deterministic sample JPEG image.
        CreateSampleJpeg(sampleImagePath);

        // 2. Build a source document that contains the sample JPEG image multiple times.
        CreateSourceDocument(sampleImagePath, sourceDocPath);

        // 3. Load the source document.
        Document srcDoc = new Document(sourceDocPath);

        // 4. Prepare a new document where blurred images will be re‑embedded.
        Document outDoc = new Document();
        DocumentBuilder outBuilder = new DocumentBuilder(outDoc);

        // 5. Iterate over all shapes that contain images.
        NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images as required.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // 5a. Extract the image bytes into a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // 5b. Load the image into an Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // 5c. Apply a simple box blur.
                    using (Bitmap blurred = ApplyBoxBlur(bitmap))
                    {
                        // 5d. Save the blurred bitmap to a new memory stream (JPEG format).
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurred.Save(blurredStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // 5e. Insert the blurred image into the output document.
                            byte[] blurredBytes = blurredStream.ToArray();
                            outBuilder.InsertImage(blurredBytes);
                            outBuilder.Writeln(); // separate images with a line break
                        }
                    }
                }
            }
        }

        // 6. Save the output document.
        outDoc.Save(outputDocPath, SaveFormat.Docx);

        // 7. Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        // Cleanup (optional): delete temporary files if desired.
    }

    // Creates a simple JPEG image with deterministic content.
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            g.Clear(Aspose.Drawing.Color.White);

            // Draw a blue rectangle.
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
            {
                g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
            }

            // Save as JPEG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

    // Builds a source document that inserts the sample image three times.
    private static void CreateSourceDocument(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(imagePath);
            builder.Writeln(); // separate images with a line break
        }

        doc.Save(docPath, SaveFormat.Docx);
    }

    // Applies a simple 3x3 box blur to the provided bitmap and returns a new blurred bitmap.
    private static Bitmap ApplyBoxBlur(Bitmap source)
    {
        int width = source.Width;
        int height = source.Height;
        Bitmap blurred = new Bitmap(width, height);

        // Iterate over each pixel.
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int sumA = 0, sumR = 0, sumG = 0, sumB = 0;
                int count = 0;

                // Accumulate colors from the 3x3 neighborhood.
                for (int ky = -1; ky <= 1; ky++)
                {
                    int ny = y + ky;
                    if (ny < 0 || ny >= height) continue;

                    for (int kx = -1; kx <= 1; kx++)
                    {
                        int nx = x + kx;
                        if (nx < 0 || nx >= width) continue;

                        Aspose.Drawing.Color pixel = source.GetPixel(nx, ny);
                        sumA += pixel.A;
                        sumR += pixel.R;
                        sumG += pixel.G;
                        sumB += pixel.B;
                        count++;
                    }
                }

                // Compute average.
                Aspose.Drawing.Color avg = Aspose.Drawing.Color.FromArgb(
                    sumA / count,
                    sumR / count,
                    sumG / count,
                    sumB / count);

                blurred.SetPixel(x, y, avg);
            }
        }

        return blurred;
    }
}
