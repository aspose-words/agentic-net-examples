using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ApplyVignetteToJpegImages
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string outputFolder = "OutputImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample JPEG image
        string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create sample Word documents containing the JPEG image
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.docx"), sampleImagePath);

        int processedCount = 0;

        // Process each document in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
                {
                    // Extract the JPEG image to a memory stream
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0;

                        // Load the image into a bitmap
                        using (Bitmap bitmap = new Bitmap(imageStream))
                        {
                            // Apply vignette effect
                            ApplyVignette(bitmap);

                            // Save the processed image
                            string outputImagePath = Path.Combine(
                                outputFolder,
                                $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.jpg");

                            bitmap.Save(outputImagePath, ImageFormat.Jpeg);
                            processedCount++;
                        }
                    }

                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was processed
        if (processedCount == 0)
            throw new InvalidOperationException("No JPEG images were found and processed.");

        Console.WriteLine($"Processing complete. {processedCount} image(s) saved to '{outputFolder}'.");
    }

    // Creates a deterministic sample JPEG image
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red circle in the center
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    int radius = 60;
                    g.FillEllipse(brush, (width - radius) / 2, (height - radius) / 2, radius, radius);
                }
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document and inserts the specified image
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple vignette effect to the bitmap
    private static void ApplyVignette(Bitmap bitmap)
    {
        int w = bitmap.Width;
        int h = bitmap.Height;
        double cx = w / 2.0;
        double cy = h / 2.0;
        double maxDist = Math.Sqrt(cx * cx + cy * cy);

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                double dx = x - cx;
                double dy = y - cy;
                double dist = Math.Sqrt(dx * dx + dy * dy);
                double factor = 1.0 - (dist / maxDist);
                factor = Math.Max(0, factor);
                factor = Math.Pow(factor, 2); // stronger vignette

                Color original = bitmap.GetPixel(x, y);
                int r = (int)(original.R * factor);
                int g = (int)(original.G * factor);
                int b = (int)(original.B * factor);
                Color newColor = Color.FromArgb(r, g, b);
                bitmap.SetPixel(x, y, newColor);
            }
        }
    }
}
