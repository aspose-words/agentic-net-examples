using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Simple record to simulate a database entry.
    public class ImageRecord
    {
        public int Id { get; set; }
        public string ImagePath { get; set; }
    }

    public static void Main()
    {
        // Directories for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample image (PNG) using Aspose.Drawing.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Build a DOCX with a Custom XML Part that contains the image as Base64.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithCustomXml(docPath, sampleImagePath);

        // 3. Load the document and extract images from its Custom XML Parts.
        List<ImageRecord> records = ExtractImagesFromCustomXml(docPath, artifactsDir);

        // 4. Serialize the mapping to JSON (simulating a DB write) and write to file.
        string jsonPath = Path.Combine(artifactsDir, "mapping.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(records, Newtonsoft.Json.Formatting.Indented));

        // Validation: ensure at least one image was extracted.
        if (records.Count == 0)
            throw new InvalidOperationException("No images were extracted from the custom XML data.");

        // Example output.
        Console.WriteLine($"Extracted {records.Count} image(s). Mapping saved to: {jsonPath}");
    }

    // Creates a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with a Custom XML Part that embeds the image as Base64.
    private static void CreateDocumentWithCustomXml(string docPath, string imagePath)
    {
        // Load image bytes and encode to Base64.
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        // Build simple XML containing the image.
        string xmlContent = $"<Images><Image id=\"1\">{base64Image}</Image></Images>";

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with Custom XML containing an image.");

        // Add the custom XML part.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

        // Save the document.
        doc.Save(docPath);
    }

    // Loads the document, parses Custom XML Parts, extracts images, and creates mapping records.
    private static List<ImageRecord> ExtractImagesFromCustomXml(string docPath, string artifactsDir)
    {
        Document doc = new Document(docPath);
        List<ImageRecord> records = new List<ImageRecord>();
        int imageCounter = 0;

        foreach (CustomXmlPart part in doc.CustomXmlParts)
        {
            // Convert the part's byte data to a string.
            string xmlString = Encoding.UTF8.GetString(part.Data);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlString);

            XmlNodeList imageNodes = xmlDoc.SelectNodes("//Image");
            foreach (XmlNode imgNode in imageNodes)
            {
                string idAttr = imgNode.Attributes["id"]?.Value ?? (imageCounter + 1).ToString();
                string base64Data = imgNode.InnerText.Trim();

                if (string.IsNullOrEmpty(base64Data))
                    continue;

                byte[] imgBytes = Convert.FromBase64String(base64Data);
                using (MemoryStream ms = new MemoryStream(imgBytes))
                {
                    // Ensure stream position is at start.
                    ms.Position = 0;

                    // Determine file name.
                    string imageFileName = $"extracted_{idAttr}.png";
                    string imageFullPath = Path.Combine(artifactsDir, imageFileName);

                    // Save the image to disk using Aspose.Drawing.
                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        bitmap.Save(imageFullPath, ImageFormat.Png);
                    }

                    // Create a record linking the image to a simulated DB entry.
                    records.Add(new ImageRecord
                    {
                        Id = int.Parse(idAttr),
                        ImagePath = imageFullPath
                    });
                }

                imageCounter++;
            }
        }

        return records;
    }
}
