using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string inputImagePath = "input.png";
        const string docPath = "DocumentWithImage.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;

                // Load the image into an Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    // Invert colors pixel by pixel.
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            Color original = bitmap.GetPixel(x, y);
                            Color inverted = Color.FromArgb(
                                255 - original.R,
                                255 - original.G,
                                255 - original.B);
                            bitmap.SetPixel(x, y, inverted);
                        }
                    }

                    // Save the inverted image to a deterministic file name.
                    string outFileName = $"extracted_{imageIndex}_inverted.png";
                    bitmap.Save(outFileName, ImageFormat.Png);
                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one image was processed.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // The program finishes automatically.
    }
}
