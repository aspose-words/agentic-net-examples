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
        // Base directory for all generated files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedDir = Path.Combine(baseDir, "Extracted");
        string watermarkedDir = Path.Combine(baseDir, "Watermarked");

        // Ensure directories exist.
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(watermarkedDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image that will be inserted into the docs.
        // -----------------------------------------------------------------
        string samplePngPath = Path.Combine(baseDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle for visual reference.
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, 160, 160);
                }
            }
            bitmap.Save(samplePngPath, ImageFormat.Png);
        }

        // ---------------------------------------------------------------
        // 2. Create a few sample Word documents each containing the PNG.
        // ---------------------------------------------------------------
        int docCount = 2;
        for (int i = 0; i < docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(samplePngPath);
            string docPath = Path.Combine(inputDocsDir, $"Doc{i + 1}.docx");
            doc.Save(docPath);
        }

        // ---------------------------------------------------------------
        // 3. Extract PNG images from each document, apply watermark, save.
        // ---------------------------------------------------------------
        int watermarkedCount = 0;
        string[] docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        for (int docIndex = 0; docIndex < docFiles.Length; docIndex++)
        {
            Document doc = new Document(docFiles[docIndex]);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Process only PNG images.
                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // -----------------------------------------------------------------
                // 3a. Save the extracted PNG to a temporary file.
                // -----------------------------------------------------------------
                string extractedPath = Path.Combine(
                    extractedDir,
                    $"extracted_doc{docIndex + 1}_img{imageIndex + 1}.png");
                shape.ImageData.Save(extractedPath);

                // -----------------------------------------------------------------
                // 3b. Load the PNG, draw watermark text, and save the result.
                // -----------------------------------------------------------------
                using (Bitmap bmp = new Bitmap(extractedPath))
                {
                    using (Graphics gfx = Graphics.FromImage(bmp))
                    {
                        // Semi‑transparent black brush.
                        using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(128, Aspose.Drawing.Color.Black)))
                        {
                            // Simple Arial font, size 24.
                            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                            {
                                // Position watermark near the bottom‑right corner.
                                float x = bmp.Width - 150;
                                float y = bmp.Height - 40;
                                gfx.DrawString("Watermark", font, brush, new PointF(x, y));
                            }
                        }
                    }

                    string watermarkedPath = Path.Combine(
                        watermarkedDir,
                        $"watermarked_doc{docIndex + 1}_img{imageIndex + 1}.png");
                    bmp.Save(watermarkedPath, ImageFormat.Png);
                    watermarkedCount++;
                }

                imageIndex++;
            }
        }

        // ---------------------------------------------------------------
        // 4. Validation – ensure at least one watermarked image was created.
        // ---------------------------------------------------------------
        if (watermarkedCount == 0)
            throw new InvalidOperationException("No PNG images were found to watermark.");

        Console.WriteLine($"Watermarked {watermarkedCount} PNG image(s). Output folder: {watermarkedDir}");
    }
}
