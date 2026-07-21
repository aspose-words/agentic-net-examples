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
        // Set up deterministic folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images with a known DPI.
        const int imageCount = 3;
        const int width = 200;
        const int height = 200;
        const float dpi = 300f;

        for (int i = 0; i < imageCount; i++)
        {
            string tiffPath = Path.Combine(inputDir, $"Sample_{i}.tiff");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                }
                bitmap.SetResolution(dpi, dpi);
                bitmap.Save(tiffPath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
            }
        }

        // Build a Word document and insert the TIFF images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string tiffFile in Directory.GetFiles(inputDir, "*.tiff"))
        {
            builder.InsertParagraph();
            builder.InsertImage(tiffFile);
        }

        // Save the document (optional, just to demonstrate loading later).
        string docPath = Path.Combine(baseDir, "SampleDoc.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // Load the document and extract all images (TIFF or otherwise).
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage)
                              .ToList();

        int extractedCount = 0;
        foreach (var shape in shapes)
        {
            // Save the image data to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing to keep metadata (e.g., DPI).
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(ms))
                {
                    // Save as lossless PNG while preserving DPI.
                    string pngPath = Path.Combine(outputDir, $"Extracted_{extractedCount}.png");
                    img.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                    extractedCount++;
                }
            }
        }

        // Validation: at least one PNG should have been created.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted and converted.");

        // Optional cleanup (commented out for inspection).
        // Directory.Delete(inputDir, true);
        // File.Delete(docPath);
    }
}
