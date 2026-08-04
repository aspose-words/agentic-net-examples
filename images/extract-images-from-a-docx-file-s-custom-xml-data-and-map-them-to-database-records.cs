using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;

namespace ExtractImagesFromCustomXml
{
    // Simple record to simulate a database entry.
    public class ImageRecord
    {
        public int Id { get; set; }
        public string ImagePath { get; set; }
    }

    public class Program
    {
        // Deterministic file names.
        private const string SampleImagePath = "sample.png";
        private const string DocumentPath = "CustomXmlImages.docx";

        public static void Main()
        {
            // 1. Create a sample image file.
            CreateSampleImage();

            // 2. Embed the image (as Base64) into a custom XML part and save the document.
            CreateDocumentWithCustomXml();

            // 3. Load the document and extract images from its custom XML parts.
            List<ImageRecord> records = ExtractImagesAndMapToRecords();

            // 4. Validate that at least one image was extracted.
            if (records.Count == 0)
                throw new InvalidOperationException("No images were extracted from the custom XML data.");

            // 5. Output the mapping (simulating database insertion).
            foreach (var rec in records)
                Console.WriteLine($"Record Id={rec.Id}, ImagePath={rec.ImagePath}");
        }

        private static void CreateSampleImage()
        {
            // Create a 100x100 white bitmap.
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction.
                graphics.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 2), 10, 10, 80, 80);
                bitmap.Save(SampleImagePath);
            }

            // Ensure the file exists.
            if (!File.Exists(SampleImagePath))
                throw new FileNotFoundException("Failed to create the sample image.", SampleImagePath);
        }

        private static void CreateDocumentWithCustomXml()
        {
            // Load the sample image bytes.
            byte[] imageBytes = File.ReadAllBytes(SampleImagePath);
            string base64Image = Convert.ToBase64String(imageBytes);

            // Build a simple XML containing the image data.
            string xmlContent = $"<root><image>{base64Image}</image></root>";

            // Create a new empty document.
            Document doc = new Document();

            // Add the custom XML part.
            string partId = Guid.NewGuid().ToString("B");
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

            // Save the document.
            doc.Save(DocumentPath);

            // Verify the document was saved.
            if (!File.Exists(DocumentPath))
                throw new FileNotFoundException("Failed to save the document with custom XML.", DocumentPath);
        }

        private static List<ImageRecord> ExtractImagesAndMapToRecords()
        {
            // Load the document that contains the custom XML part.
            Document doc = new Document(DocumentPath);

            List<ImageRecord> records = new List<ImageRecord>();
            int imageIndex = 0;

            // Iterate over all custom XML parts.
            foreach (CustomXmlPart part in doc.CustomXmlParts)
            {
                // Parse the XML content.
                XmlDocument xmlDoc = new XmlDocument();
                using (MemoryStream ms = new MemoryStream(part.Data))
                {
                    ms.Position = 0;
                    xmlDoc.Load(ms);
                }

                // Select all <image> nodes.
                XmlNodeList imageNodes = xmlDoc.GetElementsByTagName("image");
                foreach (XmlNode node in imageNodes)
                {
                    string base64 = node.InnerText.Trim();
                    if (string.IsNullOrEmpty(base64))
                        continue;

                    byte[] imgBytes = Convert.FromBase64String(base64);
                    string extractedPath = $"extracted_{imageIndex}.png";

                    // Save the extracted image.
                    using (MemoryStream imgStream = new MemoryStream(imgBytes))
                    using (FileStream fileStream = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
                    {
                        imgStream.Position = 0;
                        imgStream.CopyTo(fileStream);
                    }

                    // Verify the image file was created.
                    if (!File.Exists(extractedPath))
                        throw new FileNotFoundException("Failed to write extracted image file.", extractedPath);

                    // Simulate a database record.
                    records.Add(new ImageRecord
                    {
                        Id = imageIndex + 1,
                        ImagePath = Path.GetFullPath(extractedPath)
                    });

                    imageIndex++;
                }
            }

            return records;
        }
    }
}
