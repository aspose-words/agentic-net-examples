using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToPngConverter
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images with specific DPI
        for (int i = 1; i <= 2; i++)
        {
            string tiffPath = Path.Combine(inputDir, $"sample{i}.tif");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    // Simple visual content
                    g.DrawString($"Img {i}",
                        new Aspose.Drawing.Font("Arial", 20),
                        new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                        new Aspose.Drawing.PointF(20, 80));
                }
                // Set DPI (e.g., 150)
                bitmap.SetResolution(150f, 150f);
                bitmap.Save(tiffPath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
            }
        }

        // Create a Word document and insert the TIFF images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string tiffFile in Directory.GetFiles(inputDir, "*.tif"))
        {
            builder.InsertParagraph();
            builder.InsertImage(tiffFile);
        }
        string docPath = Path.Combine(baseDir, "DocumentWithTiffs.docx");
        doc.Save(docPath);

        // Load the document and extract images, converting them to lossless PNG while preserving DPI
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            // Save image bytes to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading

                // Load the image using Aspose.Drawing
                using (Aspose.Drawing.Image image = Aspose.Drawing.Image.FromStream(ms))
                {
                    // Prepare output file name
                    string pngPath = Path.Combine(outputDir, $"image_{imageIndex}.png");

                    // Save as PNG (lossless). DPI information is retained by the Image object.
                    image.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                }
            }
            imageIndex++;
        }

        // Validate that PNG files were created
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG files were created during conversion.");

        Console.WriteLine($"Converted {pngCount} image(s) to PNG. Output folder: {outputDir}");
    }
}
