using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create deterministic BMP sample images using Aspose.Drawing.
        // ------------------------------------------------------------
        const string sampleBmpPath = "sample.bmp";
        const int sampleWidth = 800;
        const int sampleHeight = 600;

        // First sample bitmap.
        using (Bitmap sampleBitmap = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics sampleGraphics = Graphics.FromImage(sampleBitmap))
        {
            sampleGraphics.Clear(Aspose.Drawing.Color.White);
            sampleGraphics.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Blue), 100, 100, 600, 400);
            sampleBitmap.Save(sampleBmpPath, ImageFormat.Bmp);
        }

        const string sampleBmpPath2 = "sample2.bmp";
        const int sampleWidth2 = 1200;
        const int sampleHeight2 = 900;

        // Second sample bitmap.
        using (Bitmap sampleBitmap2 = new Bitmap(sampleWidth2, sampleHeight2))
        using (Graphics sampleGraphics2 = Graphics.FromImage(sampleBitmap2))
        {
            sampleGraphics2.Clear(Aspose.Drawing.Color.White);
            sampleGraphics2.FillEllipse(new SolidBrush(Aspose.Drawing.Color.Red), 200, 150, 800, 600);
            sampleBitmap2.Save(sampleBmpPath2, ImageFormat.Bmp);
        }

        // ------------------------------------------------------------
        // Insert the BMP images into a document.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        builder.InsertParagraph();
        builder.InsertImage(sampleBmpPath2);

        const string inputDocPath = "input.docx";
        doc.Save(inputDocPath);

        // ------------------------------------------------------------
        // Load the document and resize each image to a width of 1024 px.
        // ------------------------------------------------------------
        Document loadDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        const int targetWidth = 1024;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage)
                continue;

            // Original pixel dimensions.
            int originalPixelWidth = shape.ImageData.ImageSize.WidthPixels;
            int originalPixelHeight = shape.ImageData.ImageSize.HeightPixels;

            // Guard against zero dimensions (should not happen for valid images).
            if (originalPixelWidth == 0 || originalPixelHeight == 0)
                continue;

            double scaleFactor = (double)targetWidth / originalPixelWidth;
            int targetHeight = (int)Math.Round(originalPixelHeight * scaleFactor);

            // Load original image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            using (MemoryStream originalStream = new MemoryStream(imageBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
            {
                // High‑quality resize.
                graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);

                // Save resized bitmap to a deterministic file (validation purpose).
                string resizedImagePath = $"resized_{imageIndex}.bmp";
                resizedBitmap.Save(resizedImagePath, ImageFormat.Bmp);

                // Replace the shape's image with the resized one.
                using (MemoryStream resizedStream = new MemoryStream())
                {
                    resizedBitmap.Save(resizedStream, ImageFormat.Bmp);
                    resizedStream.Position = 0; // Reset before reading.
                    shape.ImageData.SetImage(resizedStream);
                }
            }

            imageIndex++;
        }

        // ------------------------------------------------------------
        // Save the modified document.
        // ------------------------------------------------------------
        const string outputDocPath = "output.docx";
        loadDoc.Save(outputDocPath);

        // ------------------------------------------------------------
        // Validation: ensure at least one resized image file was created.
        // ------------------------------------------------------------
        string[] resizedFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "resized_*.bmp");
        if (resizedFiles.Length == 0)
            throw new InvalidOperationException("No resized BMP images were generated.");

        // ------------------------------------------------------------
        // Cleanup temporary sample files.
        // ------------------------------------------------------------
        File.Delete(sampleBmpPath);
        File.Delete(sampleBmpPath2);
    }
}
