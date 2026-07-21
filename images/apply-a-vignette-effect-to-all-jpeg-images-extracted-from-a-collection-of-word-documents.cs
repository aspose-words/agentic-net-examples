using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputDocs";
        string extractedFolder = "ExtractedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(extractedFolder);

        // Create a deterministic sample JPEG image.
        string sampleJpegPath = "sample.jpg";
        CreateSampleJpeg(sampleJpegPath, 200, 200, Aspose.Drawing.Color.LightBlue);

        // Create a few sample Word documents that contain the JPEG image.
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleJpegPath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleJpegPath);

        int totalProcessedImages = 0;

        // Process each document in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                            .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Extract original JPEG bytes.
                byte[] originalBytes = shape.ImageData.ToByteArray();

                // Load the image into Aspose.Drawing.Bitmap.
                using (MemoryStream msIn = new MemoryStream(originalBytes))
                {
                    msIn.Position = 0;
                    using (Bitmap bitmap = new Bitmap(msIn))
                    {
                        // Apply vignette effect.
                        ApplyVignette(bitmap);

                        // Save the modified image to a deterministic file.
                        string vignetteFileName = $"vignette_{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}.jpg";
                        string vignetteFullPath = Path.Combine(extractedFolder, vignetteFileName);
                        using (MemoryStream msOut = new MemoryStream())
                        {
                            bitmap.Save(msOut, ImageFormat.Jpeg);
                            msOut.Position = 0;
                            File.WriteAllBytes(vignetteFullPath, msOut.ToArray());

                            // Replace the image in the document with the vignetted version.
                            shape.ImageData.SetImage(msOut);
                        }

                        imageIndex++;
                        totalProcessedImages++;
                    }
                }
            }

            // Save the modified document.
            string outputDocPath = Path.Combine(outputFolder, $"Processed_{Path.GetFileName(docPath)}");
            doc.Save(outputDocPath);
        }

        // Validation: ensure at least one image was processed.
        if (totalProcessedImages == 0)
            throw new InvalidOperationException("No JPEG images were found and processed.");

        // Cleanup sample JPEG file.
        if (File.Exists(sampleJpegPath))
            File.Delete(sampleJpegPath);
    }

    // Creates a simple JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(fillColor);
            // Draw a simple ellipse to make the image a bit more interesting.
            using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
            {
                g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Jpeg);
                ms.Position = 0;
                File.WriteAllBytes(filePath, ms.ToArray());
            }
        }
    }

    // Creates a Word document that contains the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a vignette effect to the provided bitmap.
    private static void ApplyVignette(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;
        double centerX = width / 2.0;
        double centerY = height / 2.0;
        double maxDist = Math.Sqrt(centerX * centerX + centerY * centerY);

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                double dx = x - centerX;
                double dy = y - centerY;
                double distance = Math.Sqrt(dx * dx + dy * dy);
                // Compute a factor that darkens pixels farther from the centre.
                double factor = 1.0 - (distance / maxDist);
                factor = Math.Max(0, factor); // Clamp to [0,1]

                int r = (int)(original.R * factor);
                int g = (int)(original.G * factor);
                int b = (int)(original.B * factor);
                Aspose.Drawing.Color newColor = Aspose.Drawing.Color.FromArgb(original.A, r, g, b);
                bitmap.SetPixel(x, y, newColor);
            }
        }
    }
}
