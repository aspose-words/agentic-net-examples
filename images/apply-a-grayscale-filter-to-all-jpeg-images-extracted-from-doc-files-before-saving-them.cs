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
        // Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 10, 10, 80, 80);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Load the document and process JPEG images.
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            ImageData imgData = shape.ImageData;
            if (imgData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                imgData.Save(ms);
                ms.Position = 0; // Reset position before reading.

                // Load bitmap from stream.
                using (Bitmap bmp = new Bitmap(ms))
                {
                    // Apply grayscale filter.
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            Color pixel = bmp.GetPixel(x, y);
                            int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                            Color grayColor = Color.FromArgb(gray, gray, gray);
                            bmp.SetPixel(x, y, grayColor);
                        }
                    }

                    // Save the processed image.
                    string outputImagePath = $"extracted-{imageIndex}.jpg";
                    bmp.Save(outputImagePath, ImageFormat.Jpeg);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one grayscale image was saved.
        string[] outputFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "extracted-*.jpg");
        if (outputFiles.Length == 0)
            throw new InvalidOperationException("No JPEG images were extracted and processed.");

        // Cleanup: optional removal of intermediate files can be added here.
    }
}
