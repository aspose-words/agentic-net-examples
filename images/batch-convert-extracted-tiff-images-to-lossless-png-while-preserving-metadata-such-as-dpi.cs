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
        // Directories for sample input TIFFs and output PNGs.
        string inputDir = "InputImages";
        string outputDir = "OutputImages";
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic sample TIFF images with DPI metadata.
        for (int i = 0; i < 2; i++)
        {
            string tiffPath = Path.Combine(inputDir, $"sample{i}.tiff");
            using (Bitmap bitmap = new Bitmap(200, 100))
            {
                // Set DPI (e.g., 150).
                bitmap.SetResolution(150f, 150f);

                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    // Simple visual content.
                    g.DrawRectangle(Pens.Black, 10, 10, 180, 80);
                }

                // Save as TIFF (lossless).
                bitmap.Save(tiffPath, ImageFormat.Tiff);
            }
        }

        // Create a Word document and insert the sample TIFF images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string file in Directory.GetFiles(inputDir, "*.tiff"))
        {
            builder.InsertImage(file);
            builder.Writeln(); // Separate images.
        }

        string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Load the document and batch convert extracted images to PNG.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Get raw image bytes.
                byte[] imageBytes = shape.ImageData.ToByteArray();

                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    ms.Position = 0; // Ensure stream is at the beginning.

                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        // Preserve DPI metadata (already present in bitmap).
                        string pngPath = Path.Combine(outputDir, $"image{imageIndex}.png");
                        bitmap.Save(pngPath, ImageFormat.Png);
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one PNG was created.
        if (!Directory.GetFiles(outputDir, "*.png").Any())
            throw new InvalidOperationException("No PNG images were produced.");

        // Optional: clean up created files (comment out if inspection is needed).
        // File.Delete(docPath);
        // foreach (var f in Directory.GetFiles(inputDir)) File.Delete(f);
        // foreach (var f in Directory.GetFiles(outputDir)) File.Delete(f);
    }
}
