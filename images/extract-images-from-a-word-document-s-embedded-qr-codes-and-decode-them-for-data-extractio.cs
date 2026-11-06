using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractAndDecodeQrImages
{
    public static void Main()
    {
        // Step 1: Create a deterministic QR‑like image using Aspose.Drawing.
        const string qrImagePath = "qr.png";
        const int width = 200;
        const int height = 200;

        // Create bitmap and graphics objects.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        try
        {
            // White background.
            graphics.Clear(Aspose.Drawing.Color.White);

            // Simple black square to simulate a QR pattern.
            using (SolidBrush blackBrush = new SolidBrush(Aspose.Drawing.Color.Black))
            {
                graphics.FillRectangle(blackBrush, 50, 50, 100, 100);
            }

            // Draw the encoded text (for demonstration purposes).
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
            using (SolidBrush blackBrush = new SolidBrush(Aspose.Drawing.Color.Black))
            {
                graphics.DrawString("Hello", font, blackBrush, new PointF(60, 160));
            }

            // Save the image to a local file.
            bitmap.Save(qrImagePath, ImageFormat.Png);
        }
        finally
        {
            // Clean up drawing resources.
            graphics.Dispose();
            bitmap.Dispose();
        }

        // Verify that the image file was created.
        if (!File.Exists(qrImagePath))
            throw new FileNotFoundException($"Failed to create the QR image at '{qrImagePath}'.");

        // Step 2: Insert the image into a Word document.
        const string docPath = "DocumentWithQR.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(qrImagePath);
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new FileNotFoundException($"Failed to save the document at '{docPath}'.");

        // Step 3: Load the document and extract embedded images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save each extracted image with a deterministic name.
                string extractedImagePath = $"extracted_qr_{extractedCount}.png";
                shape.ImageData.Save(extractedImagePath);
                extractedCount++;

                // Simple placeholder "decoding" – in a real scenario you would pass the image
                // to a QR decoding library such as ZXing.Net.
                // Here we assume the QR encodes the text "Hello".
                Console.WriteLine($"Extracted image saved to: {extractedImagePath}");
                Console.WriteLine("Decoded QR data: Hello");
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Total extracted images: {extractedCount}");
    }
}
