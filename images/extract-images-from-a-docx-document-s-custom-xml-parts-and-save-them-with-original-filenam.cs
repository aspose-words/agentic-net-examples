using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // 2. Encode the image to Base64 and embed it into a custom XML part.
        string base64Image = Convert.ToBase64String(File.ReadAllBytes(sampleImagePath));
        string xmlContent = $"<images><image filename=\"{Path.GetFileName(sampleImagePath)}\">{base64Image}</image></images>";
        string customXmlPartId = Guid.NewGuid().ToString("B");

        // 3. Create a DOCX document and add the custom XML part.
        Document doc = new Document();
        doc.CustomXmlParts.Add(customXmlPartId, xmlContent);
        string docPath = Path.Combine(artifactsDir, "DocumentWithCustomXml.docx");
        doc.Save(docPath);

        // 4. Load the document and extract images from its custom XML parts.
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // Convert the part's data (byte[]) to a string.
            string partXml = Encoding.UTF8.GetString(part.Data);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(partXml);

            XmlNodeList imageNodes = xmlDoc.SelectNodes("//image");
            foreach (XmlNode imageNode in imageNodes)
            {
                string fileName = imageNode.Attributes["filename"]?.Value;
                if (string.IsNullOrEmpty(fileName))
                    continue;

                string base64Data = imageNode.InnerText;
                byte[] imageBytes = Convert.FromBase64String(base64Data);

                string outputPath = Path.Combine(artifactsDir, fileName);
                File.WriteAllBytes(outputPath, imageBytes);
                extractedCount++;
            }
        }

        // 5. Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the custom XML parts.");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Extracted {extractedCount} image(s) to '{artifactsDir}'.");
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 100x100 white bitmap.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple black rectangle.
                graphics.DrawRectangle(new Pen(Color.Black, 2), 10, 10, 80, 80);
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath);
        }
    }
}
