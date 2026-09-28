using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (200x200, light blue background).
        const string sampleImagePath = "input.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document for image extraction.
        Document loadedDoc = new Document(docPath);

        // Extract images and collect base64 data.
        List<object> extractedImages = new List<object>();
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    byte[] imageBytes = ms.ToArray();
                    string base64 = Convert.ToBase64String(imageBytes);

                    // Use the shape name if available; otherwise generate a deterministic file name.
                    string fileName = shape.Name;
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = $"image{imageIndex}.png";
                    }

                    extractedImages.Add(new { FileName = fileName, Base64 = base64 });
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
        {
            throw new Exception("No images were extracted from the document.");
        }

        // Serialize the image data to JSON.
        string json = JsonConvert.SerializeObject(extractedImages, Formatting.Indented);
        const string jsonPath = "images.json";
        File.WriteAllText(jsonPath, json);

        // Validate that the JSON file was created.
        if (!File.Exists(jsonPath))
        {
            throw new Exception("Failed to create the JSON output file.");
        }
    }
}
