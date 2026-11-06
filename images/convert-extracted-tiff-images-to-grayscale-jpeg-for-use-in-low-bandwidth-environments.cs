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
        // Step 1: Create a sample TIFF image.
        const string tiffPath = "sample.tif";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(tiffPath, ImageFormat.Tiff);
        }

        // Step 2: Create a Word document and insert the TIFF image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffPath);
        doc.Save(docPath);

        // Step 3: Load the document and extract images (process all images, the sample is TIFF).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Get image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap srcBitmap = new Bitmap(ms))
                {
                    // Convert to grayscale.
                    for (int y = 0; y < srcBitmap.Height; y++)
                    {
                        for (int x = 0; x < srcBitmap.Width; x++)
                        {
                            Color pixel = srcBitmap.GetPixel(x, y);
                            int gray = (int)(0.299 * pixel.R + 0.587 * pixel.G + 0.114 * pixel.B);
                            Color grayColor = Color.FromArgb(gray, gray, gray);
                            srcBitmap.SetPixel(x, y, grayColor);
                        }
                    }

                    // Save as JPEG.
                    string jpegPath = $"image-{imageIndex}.jpg";
                    srcBitmap.Save(jpegPath, ImageFormat.Jpeg);
                    if (!File.Exists(jpegPath))
                        throw new Exception($"Failed to create JPEG file: {jpegPath}");
                    imageIndex++;
                }
            }
        }

        // Validation: ensure at least one JPEG was created.
        if (imageIndex == 0)
            throw new Exception("No images were found and converted.");

        // Optional cleanup (commented out).
        // File.Delete(tiffPath);
        // File.Delete(docPath);
    }
}
