using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // 1. Create a deterministic sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 100);

        // 2. Encode the image to Base64 and embed it into a custom XML part
        string base64Image = Convert.ToBase64String(File.ReadAllBytes(sampleImagePath));
        string xmlContent = $"<images><image name=\"sample.png\">{base64Image}</image></images>";

        // 3. Create a new DOCX document and add the custom XML part
        Document doc = new Document();
        string partId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(partId, xmlContent);
        string docPath = Path.Combine(artifactsDir, "DocumentWithCustomXml.docx");
        doc.Save(docPath);

        // 4. Load the document (simulating a separate extraction step)
        Document loadedDoc = new Document(docPath);

        // 5. Extract images from all custom XML parts
        int extractedCount = 0;
        foreach (CustomXmlPart customPart in loadedDoc.CustomXmlParts)
        {
            // Convert the part's data (byte[]) to a UTF-8 string
            string partXml = Encoding.UTF8.GetString(customPart.Data);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(partXml);

            XmlNodeList imageNodes = xmlDoc.SelectNodes("//image");
            if (imageNodes == null) continue;

            foreach (XmlNode imageNode in imageNodes)
            {
                // Original filename is stored in the "name" attribute
                string originalFileName = imageNode.Attributes["name"]?.Value;
                if (string.IsNullOrEmpty(originalFileName)) continue;

                // Decode Base64 image data
                string base64Data = imageNode.InnerText;
                byte[] imageBytes = Convert.FromBase64String(base64Data);

                // Save the image using its original filename
                string outputPath = Path.Combine(imagesDir, originalFileName);
                using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(imageBytes, 0, imageBytes.Length);
                }

                extractedCount++;
            }
        }

        // 6. Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the custom XML parts.");

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{imagesDir}'.");
    }

    // Helper method to create a deterministic PNG image
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction
                graphics.FillRectangle(new SolidBrush(Aspose.Drawing.Color.LightBlue), 10, 10, width - 20, height - 20);
            }

            bitmap.Save(filePath);
        }
    }
}
