using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public class ImageRecord
    {
        public int Id { get; set; }
        public string ImagePath { get; set; }
    }

    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // 2. Encode the image to Base64 and embed it into custom XML
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);
        string xmlContent = $"<root><image>{base64Image}</image></root>";

        // 3. Create a DOCX document and add the custom XML part
        Document doc = new Document();
        string xmlPartId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(xmlPartId, xmlContent);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // 4. Load the document (simulating a real scenario)
        Document loadedDoc = new Document(docPath);

        // 5. Extract images from custom XML parts and map them to records
        List<ImageRecord> records = new List<ImageRecord>();
        int imageIndex = 0;

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // Convert the part data (byte[]) to a UTF-8 string
            string partXml = Encoding.UTF8.GetString(part.Data);
            XDocument xDoc = XDocument.Parse(partXml);

            foreach (XElement imgElement in xDoc.Descendants("image"))
            {
                string base64 = imgElement.Value.Trim();
                if (string.IsNullOrEmpty(base64))
                    continue;

                byte[] imgData = Convert.FromBase64String(base64);
                string extractedPath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");

                // Ensure the stream position is at the beginning before writing
                using (MemoryStream ms = new MemoryStream(imgData))
                {
                    ms.Position = 0;
                    using (FileStream fs = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
                    {
                        ms.CopyTo(fs);
                    }
                }

                records.Add(new ImageRecord
                {
                    Id = ++imageIndex,
                    ImagePath = extractedPath
                });
            }
        }

        // Validation: at least one image must be extracted
        if (records.Count == 0)
            throw new InvalidOperationException("No images were extracted from the custom XML data.");

        // 6. Output the mapping as JSON (simulating a database insert)
        string json = JsonConvert.SerializeObject(records, Formatting.Indented);
        Console.WriteLine(json);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap and fill it with a solid color
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Red);
            }

            // Save the bitmap to the specified file in PNG format
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
