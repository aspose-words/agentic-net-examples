using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image using Aspose.Drawing.
        const string sampleImagePath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 100))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    g.DrawString(
                        "Sample",
                        font,
                        new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                        new Aspose.Drawing.PointF(10, 35));
                }
            }

            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Step 2: Create a PDF document and insert the JPEG image.
        const string pdfPath = "sample.pdf";
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);
        builder.InsertImage(sampleImagePath);
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Step 3: Load the PDF and extract images with EXIF orientation correction (simulated).
        Document loadedPdf = new Document(pdfPath);
        List<string> savedImages = new List<string>();
        int imageIndex = 0;

        NodeCollection shapeNodes = loadedPdf.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0;

                    using (Aspose.Drawing.Bitmap extractedBitmap = new Aspose.Drawing.Bitmap(imageStream))
                    {
                        // Simulate EXIF orientation correction (rotate 90 degrees clockwise).
                        extractedBitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);

                        string outputPath = $"extracted_corrected_{imageIndex}.jpg";
                        extractedBitmap.Save(outputPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        savedImages.Add(outputPath);
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image was saved.
        if (savedImages.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the PDF.");
        }

        // Optional: clean up sample files (comment out if you want to keep them).
        // File.Delete(sampleImagePath);
        // File.Delete(pdfPath);
    }
}
