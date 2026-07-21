using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Simple 3x3 box blur applied to a bitmap.
    private static void ApplyBoxBlur(Aspose.Drawing.Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        // Create a copy to read original pixel values.
        using (Aspose.Drawing.Bitmap sourceCopy = new Aspose.Drawing.Bitmap(bitmap))
        {
            for (int y = 1; y < height - 1; y++)
            {
                for (int x = 1; x < width - 1; x++)
                {
                    int sumR = 0, sumG = 0, sumB = 0;

                    for (int ky = -1; ky <= 1; ky++)
                    {
                        for (int kx = -1; kx <= 1; kx++)
                        {
                            Aspose.Drawing.Color c = sourceCopy.GetPixel(x + kx, y + ky);
                            sumR += c.R;
                            sumG += c.G;
                            sumB += c.B;
                        }
                    }

                    int avgR = sumR / 9;
                    int avgG = sumG / 9;
                    int avgB = sumB / 9;

                    Aspose.Drawing.Color blurredColor = Aspose.Drawing.Color.FromArgb(avgR, avgG, avgB);
                    bitmap.SetPixel(x, y, blurredColor);
                }
            }
        }
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.jpg";
        const int imgWidth = 200;
        const int imgHeight = 150;

        // Create bitmap using Aspose.Drawing.
        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            // Obtain graphics object from the bitmap.
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a solid red rectangle.
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }

            // Save as JPEG using Aspose.Drawing.Imaging.ImageFormat.
            bmp.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Build a source document that contains the JPEG image.
        // -----------------------------------------------------------------
        const string sourceDocPath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Insert the sample image twice.
        srcBuilder.InsertImage(sampleImagePath);
        srcBuilder.InsertParagraph();
        srcBuilder.InsertImage(sampleImagePath);

        sourceDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Load the source document and process each JPEG image.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                                 .ToList();

        // Prepare a new document to hold blurred images.
        Document resultDoc = new Document();
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);

        foreach (Shape shape in shapeNodes)
        {
            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load bytes into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(ms))
                {
                    // Apply blur filter.
                    ApplyBoxBlur(bitmap);

                    // Save blurred image to a temporary stream.
                    using (MemoryStream blurredStream = new MemoryStream())
                    {
                        bitmap.Save(blurredStream, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        blurredStream.Position = 0;

                        // Insert blurred image into the result document.
                        resultBuilder.InsertImage(blurredStream);
                        resultBuilder.Writeln(); // separate images
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the resulting document containing blurred images.
        // -----------------------------------------------------------------
        const string resultDocPath = "result.docx";
        resultDoc.Save(resultDocPath);
    }
}
