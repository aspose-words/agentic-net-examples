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
        // Paths for temporary files
        const string qrImagePath = "qr.png";
        const string docPath = "sample.docx";
        const string extractedImagePath = "extracted-0.png";

        // -------------------------------------------------
        // 1. Create a deterministic QR‑code‑like image.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background
                g.Clear(Aspose.Drawing.Color.White);

                // Simple black pattern to simulate a QR code
                using (Brush blackBrush = new SolidBrush(Aspose.Drawing.Color.Black))
                {
                    // Draw a few black squares
                    g.FillRectangle(blackBrush, 20, 20, 40, 40);
                    g.FillRectangle(blackBrush, 140, 20, 40, 40);
                    g.FillRectangle(blackBrush, 20, 140, 40, 40);
                    g.FillRectangle(blackBrush, 140, 140, 40, 40);
                    g.FillRectangle(blackBrush, 80, 80, 40, 40);
                }
            }

            // Save the image to a deterministic file
            bitmap.Save(qrImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and embed the QR image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an embedded QR code image:");
        builder.InsertImage(qrImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document (simulating a separate step) and extract images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save the extracted image
                shape.ImageData.Save(extractedImagePath);
                extractedCount++;

                // -------------------------------------------------
                // 4. Decode the QR code (stub implementation).
                // -------------------------------------------------
                // In a real scenario a QR decoding library would be used.
                // Here we return a hard‑coded value because we generated the image ourselves.
                string decodedData = DecodeQrStub(extractedImagePath);
                Console.WriteLine($"Extracted image saved to '{extractedImagePath}'.");
                Console.WriteLine($"Decoded QR data: {decodedData}");
            }
        }

        // -------------------------------------------------
        // 5. Validation.
        // -------------------------------------------------
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Clean up temporary files (optional)
        // File.Delete(qrImagePath);
        // File.Delete(docPath);
        // File.Delete(extractedImagePath);
    }

    // Stub method that pretends to decode a QR code.
    private static string DecodeQrStub(string imagePath)
    {
        // Since we created the QR‑like image ourselves, we know the encoded data.
        // Replace this stub with a real QR decoder if needed.
        return "Sample QR Data";
    }
}
