using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, Pen, etc.
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string qrImagePath = "qr.png";
        const string docPath = "sample.docx";
        const string outputFolder = "ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // 1. Create a sample QR‑code‑like image.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (var bitmap = new Bitmap(imgWidth, imgHeight))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            // Fill background.
            graphics.Clear(Color.White);

            // Draw a simple black square pattern to simulate a QR code.
            using (var pen = new Pen(Color.Black, 10))
            {
                graphics.DrawRectangle(pen, 20, 20, imgWidth - 40, imgHeight - 40);
                graphics.DrawRectangle(pen, 60, 60, imgWidth - 120, imgHeight - 120);
                graphics.DrawRectangle(pen, 100, 100, imgWidth - 200, imgHeight - 200);
            }

            // Save the image to a deterministic file.
            bitmap.Save(qrImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and embed the image.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(qrImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        var loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        var extractedData = new Dictionary<string, string>(); // file name -> base64

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedFileName = Path.Combine(outputFolder, $"extracted_{imageIndex}{ext}");
                shape.ImageData.Save(extractedFileName);

                // Convert saved image to Base64.
                byte[] bytes = File.ReadAllBytes(extractedFileName);
                string base64 = Convert.ToBase64String(bytes);
                extractedData.Add(Path.GetFileName(extractedFileName), base64);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedData.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Output the extracted image data as JSON.
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(extractedData, Formatting.Indented);
        Console.WriteLine(json);
    }
}
