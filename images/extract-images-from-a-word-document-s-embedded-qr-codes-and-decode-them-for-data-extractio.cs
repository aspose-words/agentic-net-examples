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
        const string qrImagePath = "qr_SampleData.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // Step 1: Create a sample QR code image (deterministic)
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create bitmap using Aspose.Drawing
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            // Create graphics from bitmap
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // White background
                g.Clear(Aspose.Drawing.Color.White);

                // Simple black squares to simulate a QR code
                using (Aspose.Drawing.Brush blackBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.FillRectangle(blackBrush, 20, 20, 40, 40);
                    g.FillRectangle(blackBrush, 140, 20, 40, 40);
                    g.FillRectangle(blackBrush, 20, 140, 40, 40);
                    g.FillRectangle(blackBrush, 140, 140, 40, 40);
                }

                // Draw the encoded data as text (for demo decoding)
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                using (Aspose.Drawing.Brush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.DrawString("SampleData", font, textBrush, new Aspose.Drawing.PointF(50, 90));
                }
            }

            // Save the image to a deterministic file name
            bitmap.Save(qrImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a Word document and embed the QR image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with embedded QR code image:");
        builder.InsertImage(qrImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (shape.HasImage)
            {
                string extractedImagePath = $"extracted-{i}.png";
                shape.ImageData.Save(extractedImagePath);
                extractedCount++;

                // -------------------------------------------------
                // Step 4: "Decode" the QR code (demo: placeholder)
                // -------------------------------------------------
                Console.WriteLine($"Extracted image saved to: {extractedImagePath}");
                Console.WriteLine("Decoded QR data: SampleData (placeholder)");
            }
        }

        // -------------------------------------------------
        // Validation: ensure at least one image was extracted
        // -------------------------------------------------
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Clean up temporary files (optional)
        // File.Delete(qrImagePath);
        // File.Delete(docPath);
    }
}
