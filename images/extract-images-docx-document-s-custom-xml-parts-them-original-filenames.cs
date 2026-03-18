using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(sourceDocPath);

        // Iterate through all custom XML parts in the document.
        foreach (CustomXmlPart customPart in doc.CustomXmlParts)
        {
            // Convert the raw XML bytes to a string (UTF‑8 assumed).
            string xmlContent = Encoding.UTF8.GetString(customPart.Data);

            // Parse the XML.
            XDocument xDoc;
            try
            {
                xDoc = XDocument.Parse(xmlContent);
            }
            catch (Exception)
            {
                // Skip parts that are not well‑formed XML.
                continue;
            }

            // Look for <image> elements that contain a filename attribute and Base64 data.
            foreach (XElement imgElement in xDoc.Descendants("image"))
            {
                XAttribute fileNameAttr = imgElement.Attribute("filename");
                XAttribute dataAttr = imgElement.Attribute("data");

                // Some schemas may store the Base64 data as element value instead of an attribute.
                string base64Data = dataAttr != null ? dataAttr.Value : imgElement.Value;

                if (fileNameAttr == null || string.IsNullOrWhiteSpace(base64Data))
                    continue; // Required information missing.

                string fileName = fileNameAttr.Value;

                // Decode the Base64 image data.
                byte[] imageBytes;
                try
                {
                    imageBytes = Convert.FromBase64String(base64Data);
                }
                catch (FormatException)
                {
                    // Invalid Base64 string – skip this image.
                    continue;
                }

                // Build the full path and write the image file.
                string outputPath = Path.Combine(outputFolder, fileName);
                try
                {
                    File.WriteAllBytes(outputPath, imageBytes);
                    Console.WriteLine($"Saved image: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save image '{fileName}': {ex.Message}");
                }
            }
        }
    }
}
