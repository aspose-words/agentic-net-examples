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
        // Step 1: Create a sample JPEG image using Aspose.Drawing.
        const string sampleImagePath = "sample.jpg";
        const int imgWidth = 200;
        const int imgHeight = 100;

        // Create bitmap.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            // Create graphics from bitmap.
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Aspose.Drawing.Color.White);

                // Draw sample text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(10, 40));
                }
            }

            // Save the bitmap as JPEG.
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Verify that the sample image was created.
        if (!File.Exists(sampleImagePath))
            throw new InvalidOperationException($"Failed to create sample image: {sampleImagePath}");

        // Step 2: Insert the image into a Word document and save as PDF.
        const string pdfPath = "sample.pdf";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF document: {pdfPath}");

        // Step 3: Load the PDF and extract images.
        Document pdfDoc = new Document(pdfPath);
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // Reset stream position.

                // Load the image with Aspose.Drawing.
                using (Aspose.Drawing.Bitmap extractedBitmap = new Aspose.Drawing.Bitmap(imgStream))
                {
                    // Simulate reading EXIF orientation.
                    // In a real scenario, you would read the EXIF tag from the image.
                    // Here we assume an orientation value of 6 (rotate 90° clockwise).
                    int exifOrientation = 6;

                    // Apply orientation correction if needed.
                    if (exifOrientation == 3)
                        extractedBitmap.RotateFlip(RotateFlipType.Rotate180FlipNone);
                    else if (exifOrientation == 6)
                        extractedBitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    else if (exifOrientation == 8)
                        extractedBitmap.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    // Orientation value 1 (normal) requires no action.

                    // Save the corrected image.
                    string correctedImagePath = $"extracted_corrected_{imageIndex}.jpg";
                    extractedBitmap.Save(correctedImagePath, ImageFormat.Jpeg);

                    // Validate that the file was created.
                    if (!File.Exists(correctedImagePath))
                        throw new InvalidOperationException($"Failed to save corrected image: {correctedImagePath}");
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was extracted and corrected.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found in the PDF document.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(pdfPath);
    }
}
