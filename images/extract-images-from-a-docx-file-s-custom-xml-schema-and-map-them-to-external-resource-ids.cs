using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Step 2: Encode the image to Base64.
        string base64Image = Convert.ToBase64String(File.ReadAllBytes(sampleImagePath));

        // Step 3: Create a DOCX with a custom XML part that contains the image data.
        const string docPath = "sample.docx";
        CreateDocumentWithCustomXml(docPath, base64Image);

        // Step 4: Load the document and extract images from its custom XML schema.
        Document doc = new Document(docPath);
        var extractedImages = ExtractImagesFromCustomXml(doc);

        // Step 5: Map extracted images to external resource IDs.
        var resourceMap = MapImagesToExternalIds(extractedImages);

        // Step 6: Validate that at least one image was extracted and saved.
        if (resourceMap.Count == 0)
            throw new Exception("No images were extracted from the custom XML.");

        // Output the mapping (optional, just to demonstrate).
        foreach (var kvp in resourceMap)
        {
            Console.WriteLine($"{kvp.Key} -> {kvp.Value}");
        }
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 100x100 white bitmap.
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);

        // Draw a red ellipse for visual distinction.
        using (Pen pen = new Pen(Color.Red, 2))
        {
            graphics.DrawEllipse(pen, new Rectangle(10, 10, 80, 80));
        }

        // Save the bitmap to the specified path.
        bitmap.Save(path);
        graphics.Dispose();
        bitmap.Dispose();

        // Validate that the image file exists.
        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithCustomXml(string docPath, string base64Image)
    {
        // Create an empty Word document.
        Document doc = new Document();

        // Build a simple custom XML containing the image data.
        string customXml = $@"
<Images xmlns=""http://example.com/images"">
    <Image id=""img1"">{base64Image}</Image>
</Images>";

        // Add the custom XML part to the document.
        // The first argument is a unique ID for the part, the second is the XML string.
        doc.CustomXmlParts.Add("ImagesPart", customXml);

        // Save the document.
        doc.Save(docPath);

        // Validate that the document file exists.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at '{docPath}'.");
    }

    private static Dictionary<string, byte[]> ExtractImagesFromCustomXml(Document doc)
    {
        var images = new Dictionary<string, byte[]>();

        foreach (CustomXmlPart part in doc.CustomXmlParts)
        {
            // Convert the part's data (byte[]) to a UTF-8 string.
            string xmlContent = Encoding.UTF8.GetString(part.Data);
            XDocument xDoc = XDocument.Parse(xmlContent);

            // Namespace handling.
            XNamespace ns = "http://example.com/images";

            foreach (XElement imgElement in xDoc.Descendants(ns + "Image"))
            {
                XAttribute idAttr = imgElement.Attribute("id");
                if (idAttr == null) continue;

                string id = idAttr.Value;
                string base64 = imgElement.Value.Trim();

                if (string.IsNullOrEmpty(base64)) continue;

                byte[] imageBytes = Convert.FromBase64String(base64);
                string fileName = $"{id}.png";

                // Save the extracted image.
                File.WriteAllBytes(fileName, imageBytes);

                // Validate that the file was written.
                if (!File.Exists(fileName))
                    throw new Exception($"Failed to write extracted image '{fileName}'.");

                images[id] = imageBytes;
            }
        }

        return images;
    }

    private static Dictionary<string, string> MapImagesToExternalIds(Dictionary<string, byte[]> extractedImages)
    {
        var map = new Dictionary<string, string>();

        foreach (var kvp in extractedImages)
        {
            // For demonstration, generate a GUID as the external resource ID.
            string externalId = Guid.NewGuid().ToString();
            map[kvp.Key] = externalId;
        }

        return map;
    }
}
