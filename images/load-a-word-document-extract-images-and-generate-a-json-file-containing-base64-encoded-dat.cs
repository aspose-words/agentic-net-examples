using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image file (input.png).
        const string imagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
            }
            bitmap.Save(imagePath);
        }

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Step 3: Load the document for image extraction.
        Document loadedDoc = new Document(docPath);

        // Step 4: Extract images, encode them as Base64 strings.
        List<string> base64Images = new List<string>();
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
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
                    base64Images.Add(base64);
                }
            }
        }

        // Validate that at least one image was extracted.
        if (base64Images.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Step 5: Serialize the Base64 image list to JSON.
        string json = JsonConvert.SerializeObject(base64Images, Formatting.Indented);
        const string jsonPath = "images.json";
        File.WriteAllText(jsonPath, json);

        // Validate that the JSON file was created.
        if (!File.Exists(jsonPath))
        {
            throw new InvalidOperationException("Failed to create the JSON output file.");
        }
    }
}
