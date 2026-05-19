using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // 2. Encode the image as Base64 and embed it into a custom XML part
        string imageBase64 = Convert.ToBase64String(File.ReadAllBytes(sampleImagePath));
        string xmlContent = $"<images><image name=\"sample.png\">{imageBase64}</image></images>";
        string customXmlPartId = Guid.NewGuid().ToString("B");

        Document doc = new Document();
        doc.CustomXmlParts.Add(customXmlPartId, xmlContent);
        string docPath = Path.Combine(artifactsDir, "CustomXmlImages.docx");
        doc.Save(docPath);

        // 3. Load the document and extract images from its custom XML parts
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // Convert the part's data (byte[]) to a string
            string partXml = Encoding.UTF8.GetString(part.Data);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(partXml);

            XmlNodeList imageNodes = xmlDoc.SelectNodes("//image");
            if (imageNodes == null) continue;

            foreach (XmlNode imageNode in imageNodes)
            {
                if (imageNode.Attributes == null) continue;

                XmlAttribute nameAttr = imageNode.Attributes["name"];
                if (nameAttr == null) continue;

                string fileName = nameAttr.Value;
                string base64Data = imageNode.InnerText.Trim();
                if (string.IsNullOrEmpty(base64Data)) continue;

                byte[] imageBytes = Convert.FromBase64String(base64Data);
                string outputPath = Path.Combine(artifactsDir, fileName);
                File.WriteAllBytes(outputPath, imageBytes);
                extractedCount++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom XML parts.");

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{artifactsDir}'.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
