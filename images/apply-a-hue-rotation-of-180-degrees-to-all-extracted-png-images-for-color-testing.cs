using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const string inputImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill with a gradient for visual testing.
            g.Clear(Color.White);
            for (int y = 0; y < imgHeight; y++)
            {
                int red = (y * 255) / imgHeight;
                using (Brush brush = new SolidBrush(Color.FromArgb(red, 0, 255 - red)))
                {
                    g.FillRectangle(brush, 0, y, imgWidth, 1);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "Document.docx";
        doc.Save(docPath);

        // Reload the document to ensure proper image extraction.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = (Shape)shapes[i];
            if (!shape.HasImage)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap srcBitmap = new Bitmap(imgStream))
                {
                    // Apply hue rotation of 180 degrees.
                    using (Bitmap dstBitmap = new Bitmap(srcBitmap.Width, srcBitmap.Height))
                    {
                        for (int y = 0; y < srcBitmap.Height; y++)
                        {
                            for (int x = 0; x < srcBitmap.Width; x++)
                            {
                                Color srcColor = srcBitmap.GetPixel(x, y);
                                double a = srcColor.A / 255.0;
                                double r = srcColor.R / 255.0;
                                double g = srcColor.G / 255.0;
                                double b = srcColor.B / 255.0;

                                // Convert RGB to HSV.
                                double max = Math.Max(r, Math.Max(g, b));
                                double min = Math.Min(r, Math.Min(g, b));
                                double delta = max - min;

                                double h = 0;
                                if (delta != 0)
                                {
                                    if (max == r)
                                        h = 60 * (((g - b) / delta) % 6);
                                    else if (max == g)
                                        h = 60 * (((b - r) / delta) + 2);
                                    else
                                        h = 60 * (((r - g) / delta) + 4);
                                }
                                if (h < 0) h += 360;

                                double s = (max == 0) ? 0 : delta / max;
                                double v = max;

                                // Rotate hue by 180 degrees.
                                h = (h + 180) % 360;

                                // Convert HSV back to RGB.
                                double c = v * s;
                                double xVal = c * (1 - Math.Abs(((h / 60) % 2) - 1));
                                double m = v - c;

                                double r1, g1, b1;
                                if (h < 60) { r1 = c; g1 = xVal; b1 = 0; }
                                else if (h < 120) { r1 = xVal; g1 = c; b1 = 0; }
                                else if (h < 180) { r1 = 0; g1 = c; b1 = xVal; }
                                else if (h < 240) { r1 = 0; g1 = xVal; b1 = c; }
                                else if (h < 300) { r1 = xVal; g1 = 0; b1 = c; }
                                else { r1 = c; g1 = 0; b1 = xVal; }

                                int outR = (int)Math.Round((r1 + m) * 255);
                                int outG = (int)Math.Round((g1 + m) * 255);
                                int outB = (int)Math.Round((b1 + m) * 255);
                                int outA = (int)Math.Round(a * 255);

                                Color dstColor = Color.FromArgb(outA, outR, outG, outB);
                                dstBitmap.SetPixel(x, y, dstColor);
                            }
                        }

                        // Save the processed image.
                        string outputPath = $"extracted-{i + 1}.png";
                        dstBitmap.Save(outputPath);
                        extractedCount++;
                    }
                }
            }
        }

        // Validation: ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Cleanup: optional removal of intermediate files can be added here.
    }
}
