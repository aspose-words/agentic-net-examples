using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsCustomXmlImageExtractor
{
    public static class ImageExtractor
    {
        public static Dictionary<string, string> ExtractImages(string docPath, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);
            Document doc = new Document(docPath);
            var resourceMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (CustomXmlPart xmlPart in doc.CustomXmlParts)
            {
                string xmlContent = Encoding.UTF8.GetString(xmlPart.Data);
                XDocument xDoc;
                try
                {
                    xDoc = XDocument.Parse(xmlContent);
                }
                catch
                {
                    continue;
                }

                foreach (XElement imgElement in xDoc.Descendants("image"))
                {
                    XAttribute idAttr = imgElement.Attribute("id");
                    if (idAttr == null) continue;

                    string resourceId = idAttr.Value.Trim();
                    if (string.IsNullOrEmpty(resourceId)) continue;

                    string base64Data = imgElement.Value.Trim();
                    if (string.IsNullOrEmpty(base64Data)) continue;

                    byte[] imageBytes;
                    try
                    {
                        imageBytes = Convert.FromBase64String(base64Data);
                    }
                    catch (FormatException)
                    {
                        continue;
                    }

                    string extension = GetImageExtension(imageBytes) ?? ".bin";
                    string fileName = $"{resourceId}{extension}";
                    string filePath = Path.Combine(outputFolder, fileName);
                    File.WriteAllBytes(filePath, imageBytes);
                    resourceMap[resourceId] = filePath;
                }
            }

            return resourceMap;
        }

        private static string GetImageExtension(byte[] data)
        {
            if (data.Length < 4) return null;
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return ".png";
            if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return ".jpg";
            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38) return ".gif";
            if (data[0] == 0x42 && data[1] == 0x4D) return ".bmp";
            if (data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00) return ".tif";
            if (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A) return ".tif";
            return null;
        }

        private static void CreateSampleDocument(string path)
        {
            // 1x1 transparent PNG base64
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3ZcAAAAASUVORK5CYII=";
            string xml = $"<root><image id=\"img1\">{pngBase64}</image></root>";

            Document doc = new Document();
            doc.CustomXmlParts.Add("customXmlId", xml);
            doc.Save(path);
        }

        public static void Main()
        {
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeSample");
            Directory.CreateDirectory(tempFolder);

            string sourceDoc = Path.Combine(tempFolder, "SampleWithCustomXml.docx");
            string imagesFolder = Path.Combine(tempFolder, "ExtractedImages");

            // Create a sample DOCX with a custom XML part if it doesn't exist.
            if (!File.Exists(sourceDoc))
                CreateSampleDocument(sourceDoc);

            Dictionary<string, string> map = ExtractImages(sourceDoc, imagesFolder);

            Console.WriteLine("Extracted images:");
            foreach (var kvp in map)
                Console.WriteLine($"Resource ID: {kvp.Key} => File: {kvp.Value}");
        }
    }
}
