using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;          // Needed for CustomXmlPart
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string mappingCsvPath = "mapping.csv";

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing and save it locally.
        // -----------------------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Load the image bytes and encode them as Base64 for embedding.
        // -----------------------------------------------------------------
        byte[] imageBytes = File.ReadAllBytes(sampleImagePath);
        string base64Image = Convert.ToBase64String(imageBytes);

        // -----------------------------------------------------------------
        // 3. Create a new blank Word document and insert the sample image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(sampleImagePath);
        shape.Width = imgWidth;
        shape.Height = imgHeight;

        // -----------------------------------------------------------------
        // 4. Create a custom XML part that contains the Base64 image data.
        // -----------------------------------------------------------------
        string xmlContent = $@"
<root>
    <record id='1'>
        <image>{base64Image}</image>
    </record>
    <record id='2'>
        <image>{base64Image}</image>
    </record>
</root>";
        string xmlPartId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 5. Load the document back and extract images from its custom XML.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var imageMappings = new Dictionary<string, string>(); // recordId -> extracted file path

        foreach (CustomXmlPart customPart in loadedDoc.CustomXmlParts)
        {
            // Convert the part's data (byte[]) to a string.
            string partXml = Encoding.UTF8.GetString(customPart.Data);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(partXml);

            XmlNodeList recordNodes = xmlDoc.SelectNodes("//record");
            if (recordNodes == null) continue;

            foreach (XmlNode recordNode in recordNodes)
            {
                XmlAttribute idAttr = recordNode.Attributes["id"];
                if (idAttr == null) continue;

                string recordId = idAttr.Value;
                XmlNode imageNode = recordNode.SelectSingleNode("image");
                if (imageNode == null) continue;

                string base64 = imageNode.InnerText.Trim();
                if (string.IsNullOrEmpty(base64)) continue;

                // Decode Base64 to raw image bytes.
                byte[] decodedBytes = Convert.FromBase64String(base64);

                // Save the extracted image to a deterministic file name.
                string extractedImagePath = $"extracted_{recordId}.png";
                using (MemoryStream ms = new MemoryStream(decodedBytes))
                {
                    ms.Position = 0; // Ensure the stream is at the beginning.
                    using (FileStream fs = new FileStream(extractedImagePath, FileMode.Create, FileAccess.Write))
                    {
                        ms.CopyTo(fs);
                    }
                }

                // Record the mapping.
                imageMappings[recordId] = Path.GetFullPath(extractedImagePath);
            }
        }

        // -----------------------------------------------------------------
        // 6. Validation: ensure at least one image was extracted.
        // -----------------------------------------------------------------
        if (imageMappings.Count == 0)
            throw new InvalidOperationException("No images were extracted from the custom XML parts.");

        // -----------------------------------------------------------------
        // 7. Write the record-to-image mapping to a CSV file.
        // -----------------------------------------------------------------
        using (StreamWriter writer = new StreamWriter(mappingCsvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("RecordId,ImageFilePath");
            foreach (var kvp in imageMappings)
            {
                writer.WriteLine($"{kvp.Key},{kvp.Value}");
            }
        }

        // -----------------------------------------------------------------
        // 8. Final validation: ensure the CSV file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(mappingCsvPath))
            throw new FileNotFoundException("Mapping CSV file was not created.", mappingCsvPath);
    }
}
