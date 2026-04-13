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
        const string documentPath = "sample.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        // Fill background with white.
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple blue rectangle.
        using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Blue))
        {
            graphics.FillRectangle(brush, 20, 20, 160, 60);
        }
        // Save the image to a local file.
        bitmap.Save(inputImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the PNG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // -----------------------------------------------------------------
        // 3. Extract all PNG images from the document, invert their colors,
        //    and save the inverted versions.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure the stream is at the beginning.
                using (Bitmap imgBitmap = new Bitmap(ms))
                {
                    // Invert colors pixel by pixel.
                    for (int y = 0; y < imgBitmap.Height; y++)
                    {
                        for (int x = 0; x < imgBitmap.Width; x++)
                        {
                            Aspose.Drawing.Color original = imgBitmap.GetPixel(x, y);
                            Aspose.Drawing.Color inverted = Aspose.Drawing.Color.FromArgb(
                                255 - original.R,
                                255 - original.G,
                                255 - original.B);
                            imgBitmap.SetPixel(x, y, inverted);
                        }
                    }

                    // Save the inverted image.
                    string invertedPath = $"inverted_{imageIndex}.png";
                    imgBitmap.Save(invertedPath);
                    if (!File.Exists(invertedPath))
                        throw new InvalidOperationException($"Failed to create {invertedPath}.");

                    imageIndex++;
                }
            }
        }

        // Validate that at least one inverted image was produced.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to invert.");

        // Cleanup temporary files (optional).
        // File.Delete(inputImagePath);
        // File.Delete(documentPath);
    }
}
