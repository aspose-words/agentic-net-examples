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
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(inputDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill with a solid color (light blue) for visibility.
            g.Clear(Color.LightBlue);
            // Draw a simple red rectangle to have distinct colors.
            using (Brush redBrush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(redBrush, 50, 50, 100, 100);
            }
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the PNG image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(inputDir, "DocumentWithImage.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all PNG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Retrieve the image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into a Bitmap for manipulation.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // -----------------------------------------------------------------
                // 4. Apply a simple color‑balance adjustment.
                //    Here we increase the red channel by 50 (clamped to 255).
                // -----------------------------------------------------------------
                for (int y = 0; y < bitmap.Height; y++)
                {
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        Color original = bitmap.GetPixel(x, y);
                        int newR = Math.Min(original.R + 50, 255);
                        Color adjusted = Color.FromArgb(original.A, newR, original.G, original.B);
                        bitmap.SetPixel(x, y, adjusted);
                    }
                }

                // -----------------------------------------------------------------
                // 5. Save the adjusted image to the output folder.
                // -----------------------------------------------------------------
                string outFile = Path.Combine(outputDir, $"extracted_{extractedCount}.png");
                bitmap.Save(outFile);
                extractedCount++;
            }
        }

        // -----------------------------------------------------------------
        // 6. Validation – ensure at least one image was written.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Optional: clean up temporary files (comment out if you want to inspect them).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
