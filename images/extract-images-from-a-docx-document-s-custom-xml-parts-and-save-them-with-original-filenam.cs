using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImageFile = "sample.png";
        CreateSampleImage(sampleImageFile);

        // Step 2: Create a DOCX document that contains a custom XML part with the image data.
        const string docFile = "CustomXmlImages.docx";
        CreateDocumentWithCustomXml(docFile, sampleImageFile);

        // Step 3: Extract images from the custom XML parts and save them with their original filenames.
        ExtractImagesFromCustomXml(docFile);
    }

    // Creates a simple 100x100 white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        var bitmap = new Bitmap(100, 100);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Builds a document, adds a custom XML part that stores the image as Base64,
    // and saves the document to the specified path.
    private static void CreateDocumentWithCustomXml(string docPath, string imagePath)
    {
        var doc = new Document();

        // Read the image bytes and encode them as Base64.
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        string base64 = Convert.ToBase64String(imageBytes);

        // Simple XML that holds the image name and its Base64 data.
        string xmlContent = $"<root><image name=\"{Path.GetFileName(imagePath)}\">{base64}</image></root>";

        // Add the custom XML part to the document.
        string partId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(partId, xmlContent);

        // Save the document.
        doc.Save(docPath);
    }

    // Loads the document, parses each custom XML part, extracts image data,
    // and saves each image using Shape/ImageData APIs.
    private static void ExtractImagesFromCustomXml(string docPath)
    {
        var doc = new Document(docPath);
        int extractedCount = 0;

        foreach (CustomXmlPart part in doc.CustomXmlParts)
        {
            // Convert the part's raw data (byte[]) to a UTF‑8 string.
            string xmlString = Encoding.UTF8.GetString(part.Data);
            var xdoc = XDocument.Parse(xmlString);

            // Find all <image> elements.
            var imageElements = xdoc.Descendants("image");
            foreach (var imgElem in imageElements)
            {
                string fileName = (string)imgElem.Attribute("name");
                string base64 = imgElem.Value.Trim();

                if (string.IsNullOrEmpty(base64) || string.IsNullOrEmpty(fileName))
                    continue;

                // Decode the Base64 image data.
                byte[] imgBytes = Convert.FromBase64String(base64);

                // Use a temporary document and DocumentBuilder to insert the image,
                // which automatically creates a Shape with the image data.
                using (var ms = new MemoryStream(imgBytes))
                {
                    ms.Position = 0; // Ensure the stream is at the beginning.

                    var tempDoc = new Document();
                    var builder = new DocumentBuilder(tempDoc);
                    Shape shape = builder.InsertImage(ms); // Insertion strict rule.

                    // Validate that the shape indeed contains an image.
                    if (!shape.HasImage)
                        throw new InvalidOperationException("Inserted shape does not contain an image.");

                    // Save the image using ImageData.Save (core image rule).
                    shape.ImageData.Save(fileName);
                    extractedCount++;
                }
            }
        }

        // Validation rule: at least one image must have been extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the custom XML parts.");
    }
}
