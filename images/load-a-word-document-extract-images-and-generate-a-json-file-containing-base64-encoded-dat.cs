using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string imagePath = "sample.png";
        const string docPath = "sample.docx";
        const string jsonPath = "images.json";

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Explicit Aspose.Drawing types as required by the image creation rules.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw deterministic text.
        graphics.DrawString(
            "Sample",
            new Aspose.Drawing.Font("Arial", 20),
            new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
            new Aspose.Drawing.PointF(10, 80));

        // Save the generated image.
        bitmap.Save(imagePath);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        // Save the document for later loading.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var extractedImages = new List<ImageInfo>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream.
            using (var ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Ensure the stream is at the beginning.
                byte[] imageBytes = ms.ToArray();
                string base64 = Convert.ToBase64String(imageBytes);
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"image{imageIndex}{extension}";
                extractedImages.Add(new ImageInfo { FileName = fileName, Base64Data = base64 });
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were found in the document.");

        // -------------------------------------------------
        // 4. Serialize the extracted images to JSON.
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(extractedImages, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // -------------------------------------------------
        // 5. Verify that the JSON file was created.
        // -------------------------------------------------
        if (!File.Exists(jsonPath))
            throw new FileNotFoundException("Failed to create the JSON output file.", jsonPath);
    }

    // Helper class representing one extracted image.
    private class ImageInfo
    {
        public string FileName { get; set; }
        public string Base64Data { get; set; }
    }
}
