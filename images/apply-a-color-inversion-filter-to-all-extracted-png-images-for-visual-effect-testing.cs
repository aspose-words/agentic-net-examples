using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string sampleImagePath = "sample.png";
        const string documentPath = "sample.docx";

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
                // Draw a red rectangle.
                g.FillRectangle(new SolidBrush(Color.Red), 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save the image to a local file.
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document (already in memory) and extract PNG images.
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Obtain raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap bmp = new Bitmap(ms))
                {
                    // Invert colors pixel by pixel.
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            Color original = bmp.GetPixel(x, y);
                            Color inverted = Color.FromArgb(
                                255 - original.R,
                                255 - original.G,
                                255 - original.B);
                            bmp.SetPixel(x, y, inverted);
                        }
                    }

                    // Save the inverted image.
                    string outFileName = $"inverted_{extractedCount}.png";
                    bmp.Save(outFileName);
                }
            }

            extractedCount++;
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was processed.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were found to process.");

        // The program finishes automatically.
    }
}
