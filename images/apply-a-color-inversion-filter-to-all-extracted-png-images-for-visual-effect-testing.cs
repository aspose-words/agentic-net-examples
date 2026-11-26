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
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            graphics.Clear(Color.White);
            // Draw a simple red rectangle.
            graphics.FillRectangle(new SolidBrush(Color.Red), 50, 50, 100, 100);
            // Save the image.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Step 2: Create a Word document and insert the sample PNG.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Extract all PNG images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Ensure the image is a PNG.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            string extractedPath = $"extracted-{extractedCount}.png";
            shape.ImageData.Save(extractedPath);
            extractedCount++;

            // Step 4: Apply color inversion to the extracted PNG.
            using (Bitmap extractedBitmap = new Bitmap(extractedPath))
            {
                for (int y = 0; y < extractedBitmap.Height; y++)
                {
                    for (int x = 0; x < extractedBitmap.Width; x++)
                    {
                        Color original = extractedBitmap.GetPixel(x, y);
                        Color inverted = Color.FromArgb(
                            255 - original.R,
                            255 - original.G,
                            255 - original.B);
                        extractedBitmap.SetPixel(x, y, inverted);
                    }
                }

                string invertedPath = $"inverted-{extractedCount - 1}.png";
                extractedBitmap.Save(invertedPath, ImageFormat.Png);
            }
        }

        // Validation: ensure at least one inverted image was produced.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        bool anyInvertedExists = false;
        for (int i = 0; i < extractedCount; i++)
        {
            if (File.Exists($"inverted-{i}.png"))
            {
                anyInvertedExists = true;
                break;
            }
        }

        if (!anyInvertedExists)
            throw new InvalidOperationException("Inverted PNG images were not created.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // for (int i = 0; i < extractedCount; i++)
        // {
        //     File.Delete($"extracted-{i}.png");
        //     File.Delete($"inverted-{i}.png");
        // }
    }
}
