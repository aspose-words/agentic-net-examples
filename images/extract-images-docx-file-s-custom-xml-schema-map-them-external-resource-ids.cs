using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsCustomXmlImageExtractor
{
    /// <summary>
    /// Demonstrates how to extract images that are stored inside custom XML parts of a DOCX file
    /// and map them to external resource identifiers.
    /// </summary>
    public static class ImageExtractor
    {
        /// <summary>
        /// Loads a DOCX document, parses each custom XML part, extracts Base64‑encoded images,
        /// saves them to the specified folder and returns a map of resource IDs to file paths.
        /// </summary>
        /// <param name="docPath">Full path to the source DOCX file.</param>
        /// <param name="outputFolder">Folder where extracted images will be written.</param>
        /// <returns>Dictionary where the key is the external resource ID and the value is the saved image file path.</returns>
        public static Dictionary<string, string> ExtractImages(string docPath, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the document (lifecycle rule: load).
            Document doc = new Document(docPath);

            // Prepare the result map.
            var resourceMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Iterate over all custom XML parts in the document.
            foreach (CustomXmlPart xmlPart in doc.CustomXmlParts)
            {
                // The XML content is stored as a byte array; convert it to a string.
                string xmlContent = Encoding.UTF8.GetString(xmlPart.Data);

                // Parse the XML. The schema is user‑defined, so we look for any element named "image"
                // that contains a Base64 string and an attribute "id" that serves as the external resource ID.
                XDocument xDoc;
                try
                {
                    xDoc = XDocument.Parse(xmlContent);
                }
                catch (Exception)
                {
                    // Skip malformed XML parts.
                    continue;
                }

                // Find all <image> elements (adjust the element name if your schema differs).
                foreach (XElement imgElement in xDoc.Descendants("image"))
                {
                    // Retrieve the resource identifier.
                    XAttribute idAttr = imgElement.Attribute("id");
                    if (idAttr == null) continue; // No ID – cannot map.

                    string resourceId = idAttr.Value.Trim();
                    if (string.IsNullOrEmpty(resourceId)) continue;

                    // Retrieve the Base64 image data.
                    string base64Data = imgElement.Value.Trim();
                    if (string.IsNullOrEmpty(base64Data)) continue;

                    // Decode the image bytes.
                    byte[] imageBytes;
                    try
                    {
                        imageBytes = Convert.FromBase64String(base64Data);
                    }
                    catch (FormatException)
                    {
                        // Invalid Base64 – skip this element.
                        continue;
                    }

                    // Determine a file extension. Aspose.Words can infer the image type from the byte header.
                    // For simplicity we use the first few bytes to guess common formats.
                    string extension = GetImageExtension(imageBytes);
                    if (extension == null) extension = ".bin"; // Fallback.

                    // Build a unique file name using the resource ID.
                    string fileName = $"{resourceId}{extension}";
                    string filePath = Path.Combine(outputFolder, fileName);

                    // Save the image to disk (lifecycle rule: save).
                    File.WriteAllBytes(filePath, imageBytes);

                    // Add the mapping to the result dictionary.
                    resourceMap[resourceId] = filePath;
                }
            }

            return resourceMap;
        }

        /// <summary>
        /// Very small helper that inspects the first bytes of an image to guess its file extension.
        /// Supports PNG, JPEG, GIF, BMP and TIFF. Returns null if the format is unknown.
        /// </summary>
        private static string GetImageExtension(byte[] data)
        {
            if (data.Length < 4) return null;

            // PNG: 89 50 4E 47
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return ".png";

            // JPEG: FF D8 FF
            if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
                return ".jpg";

            // GIF: 47 49 46 38
            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38)
                return ".gif";

            // BMP: 42 4D
            if (data[0] == 0x42 && data[1] == 0x4D)
                return ".bmp";

            // TIFF (little endian): 49 49 2A 00
            if (data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00)
                return ".tif";

            // TIFF (big endian): 4D 4D 00 2A
            if (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A)
                return ".tif";

            return null;
        }

        // Example usage.
        public static void Main()
        {
            string sourceDoc = @"C:\Docs\SampleWithCustomXml.docx";
            string imagesFolder = @"C:\Docs\ExtractedImages";

            Dictionary<string, string> map = ExtractImages(sourceDoc, imagesFolder);

            Console.WriteLine("Extracted images:");
            foreach (var kvp in map)
            {
                Console.WriteLine($"Resource ID: {kvp.Key} => File: {kvp.Value}");
            }
        }
    }
}
