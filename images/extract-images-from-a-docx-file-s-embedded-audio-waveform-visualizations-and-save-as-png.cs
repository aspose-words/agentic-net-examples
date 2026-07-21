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
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will represent an
        //    audio waveform visualization.
        // -----------------------------------------------------------------
        string waveformPath = Path.Combine(artifactsDir, "waveform.png");
        const int imgWidth = 400;
        const int imgHeight = 100;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // White background.
            g.Clear(Color.White);

            // Draw a simple waveform (sine‑like line).
            Pen pen = new Pen(Color.Blue, 2);
            for (int x = 0; x < imgWidth; x++)
            {
                int y = imgHeight / 2 + (int)(30 * Math.Sin(2 * Math.PI * x / 50));
                if (x == 0)
                    g.DrawLine(pen, x, y, x, y);
                else
                    g.DrawLine(pen, x - 1, imgHeight / 2 + (int)(30 * Math.Sin(2 * Math.PI * (x - 1) / 50)), x, y);
            }

            // Save the image to disk.
            bitmap.Save(waveformPath);
        }

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and embed the sample waveform image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as an inline shape.
        builder.InsertImage(waveformPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images (including the
        //    waveform visualization) to PNG files.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine output file name (force PNG extension).
                string outFile = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");

                // Save the image data to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;

                    // Load the image via Aspose.Drawing and re‑save as PNG.
                    using (Bitmap img = new Bitmap(ms))
                    {
                        img.Save(outFile, ImageFormat.Png);
                    }
                }

                extractedCount++;
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
