using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image (PNG) using Aspose.Drawing
        const string imagePath = "sample.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
        {
            graphics.DrawRectangle(pen, 10, 10, 80, 80);
        }
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Create a DOCX document and insert the image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Step 3: Reload the document and extract the first image as BMP into a MemoryStream
        Document loadedDoc = new Document(docPath);
        MemoryStream bmpStream = null;

        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                // Save the original image (PNG) to a temporary stream
                using (MemoryStream pngStream = new MemoryStream())
                {
                    shape.ImageData.Save(pngStream);
                    pngStream.Position = 0;

                    // Load the PNG with Aspose.Drawing and re‑save as BMP
                    using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(pngStream))
                    {
                        bmpStream = new MemoryStream();
                        img.Save(bmpStream, ImageFormat.Bmp);
                        bmpStream.Position = 0; // Reset for reading
                    }
                }
                break; // Extract only the first image for this example
            }
        }

        if (bmpStream == null || bmpStream.Length == 0)
            throw new InvalidOperationException("No image was extracted from the document.");

        // Step 4: Simulate passing the BMP stream to an API by creating a JSON payload
        string base64Image = Convert.ToBase64String(bmpStream.ToArray());
        var payload = new { ImageBase64 = base64Image };
        string json = JsonConvert.SerializeObject(payload, Formatting.Indented);
        const string jsonPath = "payload.json";
        File.WriteAllText(jsonPath, json);

        // Validation: ensure the JSON file was created and contains data
        if (!File.Exists(jsonPath) || new FileInfo(jsonPath).Length == 0)
            throw new InvalidOperationException("Failed to create the JSON payload file.");

        // Clean up resources
        bmpStream.Dispose();

        // Program completed successfully.
    }
}
