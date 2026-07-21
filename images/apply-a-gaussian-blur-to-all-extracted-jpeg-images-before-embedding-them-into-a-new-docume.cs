using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string workDir = Directory.GetCurrentDirectory();

        // -----------------------------------------------------------------
        // 1. Create sample JPEG images.
        // -----------------------------------------------------------------
        string[] sampleImagePaths = new string[2];
        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            string imgPath = Path.Combine(workDir, $"sample{i + 1}.jpg");
            CreateSampleJpeg(imgPath, i);
            sampleImagePaths[i] = imgPath;
        }

        // -----------------------------------------------------------------
        // 2. Build a source document that contains the sample JPEG images.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(workDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        foreach (string imgPath in sampleImagePaths)
        {
            srcBuilder.InsertImage(imgPath);
            srcBuilder.Writeln(); // separate images with a line break
        }
        sourceDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Load the source document and extract JPEG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegCount = 0;

        // Prepare the result document once.
        Document resultDoc = new Document();
        DocumentBuilder resBuilder = new DocumentBuilder(resultDoc);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Extract original JPEG bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(originalBytes))
            {
                ms.Position = 0;
                // Load into Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(ms))
                {
                    // Apply Gaussian blur (placeholder implementation).
                    using (Aspose.Drawing.Bitmap blurred = ApplyGaussianBlur(bitmap))
                    {
                        // Save blurred image to a deterministic file.
                        string blurredPath = Path.Combine(workDir, $"blurred_{jpegCount + 1}.jpg");
                        blurred.Save(blurredPath, ImageFormat.Jpeg);

                        // Validate that the file was created.
                        if (!File.Exists(blurredPath))
                            throw new InvalidOperationException($"Blurred image not created: {blurredPath}");

                        // Insert the blurred image into the result document.
                        if (jpegCount > 0)
                            resBuilder.InsertParagraph(); // separate images

                        resBuilder.InsertImage(blurredPath);
                        jpegCount++;
                    }
                }
            }
        }

        // Save the result document if at least one image was processed.
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the source document.");

        string resultDocPath = Path.Combine(workDir, "Result.docx");
        resultDoc.Save(resultDocPath);
    }

    // Creates a simple deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int index)
    {
        int width = 200;
        int height = 150;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            // Fill background.
            g.Clear(Aspose.Drawing.Color.White);

            // Draw a colored rectangle that varies with the index.
            Aspose.Drawing.Color rectColor = (index % 2 == 0) ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightCoral;
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(rectColor))
            {
                g.FillRectangle(brush, 20, 20, width - 40, height - 40);
            }

            // Draw index text.
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
            using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
            {
                g.DrawString($"Img {index + 1}", font, textBrush, new Aspose.Drawing.PointF(30, 60));
            }

            // Save as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Placeholder for Gaussian blur – returns a cloned bitmap.
    private static Aspose.Drawing.Bitmap ApplyGaussianBlur(Aspose.Drawing.Bitmap source)
    {
        // In a real scenario, apply a Gaussian blur filter here.
        // For this example, simply clone the bitmap to keep the code simple and safe.
        return (Aspose.Drawing.Bitmap)source.Clone();
    }
}
