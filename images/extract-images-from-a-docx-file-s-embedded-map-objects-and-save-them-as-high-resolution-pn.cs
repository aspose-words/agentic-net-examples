using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class ExtractMapImages
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample high‑resolution PNG image (simulating a map)
        string sampleImagePath = Path.Combine(artifactsDir, "sample_map.png");
        const int imgWidth = 1200;
        const int imgHeight = 800;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple map‑like grid
                Pen pen = new Pen(Color.LightGray, 2);
                for (int x = 0; x <= imgWidth; x += 100)
                    g.DrawLine(pen, x, 0, x, imgHeight);
                for (int y = 0; y <= imgHeight; y += 100)
                    g.DrawLine(pen, 0, y, imgWidth, y);
                pen.Dispose();

                // Draw a red circle to represent a point of interest
                Brush brush = new SolidBrush(Color.Red);
                int radius = 30;
                g.FillEllipse(brush, imgWidth / 2 - radius, imgHeight / 2 - radius, radius * 2, radius * 2);
                brush.Dispose();
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // 2. Create a DOCX document and embed the PNG image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "input.docx");
        doc.Save(docPath);

        // 3. Load the document (simulating a separate load operation)
        Document loadedDoc = new Document(docPath);

        // 4. Extract all images from shape nodes and save them as high‑resolution PNG files
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on the image type; force PNG if not already PNG
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                if (!ext.Equals(".png", StringComparison.OrdinalIgnoreCase))
                {
                    ext = ".png";
                }

                string outFile = Path.Combine(artifactsDir, $"extracted_image_{imageIndex}{ext}");
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }

        // 5. Validation – ensure at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: output result count (no interactive prompt)
        Console.WriteLine($"Extracted {imageIndex} image(s) to folder: {artifactsDir}");
    }
}
