using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Fields;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a Word document and insert a QR‑code‑like image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a deterministic sample image using Aspose.Drawing.
        const int imgSize = 200;
        string sampleImagePath = Path.Combine(outputDir, "qr_sample.png");

        // Create bitmap, draw placeholder text, and save to file.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgSize, imgSize))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw placeholder text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    g.DrawString("HelloWorld", font, Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(10, imgSize / 2 - 10));
                }
            }

            // Save the bitmap to a deterministic file.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Insert the image file into the document.
        Shape qrShape = builder.InsertImage(sampleImagePath);
        // Store the original value for later "decoding".
        qrShape.Title = "HelloWorld";

        // Save the document containing the placeholder QR code.
        string docPath = Path.Combine(outputDir, "QrDocument.docx");
        doc.Save(docPath);

        // -------------------------------------------------
        // 2. Load the document and extract images from shapes.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        List<QrResult> results = new List<QrResult>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"QrImage_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the extracted image.
                shape.ImageData.Save(imagePath);

                // "Decode" the QR code by reading the stored Title.
                string decodedValue = shape.Title ?? string.Empty;

                results.Add(new QrResult
                {
                    Index = imageIndex,
                    ImageFile = imageFileName,
                    DecodedValue = decodedValue
                });

                Console.WriteLine($"Extracted image: {imageFileName}, decoded QR value: {decodedValue}");
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (results.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 3. Serialize extraction results to JSON.
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(results, Formatting.Indented);
        string jsonPath = Path.Combine(outputDir, "QrData.json");
        File.WriteAllText(jsonPath, json);
        Console.WriteLine($"Extraction details saved to: {jsonPath}");
    }

    // Helper class to hold extraction details.
    private class QrResult
    {
        public int Index { get; set; }
        public string ImageFile { get; set; }
        public string DecodedValue { get; set; }
    }
}
