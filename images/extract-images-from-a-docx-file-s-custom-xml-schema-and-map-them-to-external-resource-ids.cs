using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath, 200, 100);

        // Step 2: Build a DOCX containing the image and a custom XML part that maps the image to an external ID.
        const string docPath = "sample.docx";
        CreateDocumentWithCustomXml(docPath, sampleImagePath);

        // Step 3: Load the document, parse the custom XML, extract the mapped images, and save them using the external IDs.
        var mapping = ExtractImagesBasedOnCustomXml(docPath);

        // Step 4: Validate that at least one image was extracted.
        if (mapping.Count == 0)
            throw new InvalidOperationException("No images were extracted based on the custom XML mapping.");

        // Step 5: Output the mapping as JSON.
        string json = JsonConvert.SerializeObject(mapping, Formatting.Indented);
        Console.WriteLine(json);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Create a bitmap using Aspose.Drawing.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill background.
        graphics.Clear(Color.White);

        // Draw a simple rectangle.
        using (Pen pen = new Pen(Color.Blue, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the image.
        bitmap.Save(path);

        // Clean up.
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateDocumentWithCustomXml(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the sample image and assign a deterministic name.
        Shape imageShape = builder.InsertImage(imagePath);
        imageShape.Name = "MyImage1";

        // Define custom XML that maps the image name to an external resource ID.
        const string customXml = @"
<Resources xmlns=""http://example.com/resources"">
    <Resource>
        <ImageName>MyImage1</ImageName>
        <ExternalId>ext-001</ExternalId>
    </Resource>
</Resources>";

        // Add the custom XML part to the document.
        CustomXmlPart xmlPart = new CustomXmlPart();
        xmlPart.Data = Encoding.UTF8.GetBytes(customXml);
        doc.CustomXmlParts.Add(xmlPart);

        // Save the document.
        doc.Save(docPath);
    }

    private static Dictionary<string, string> ExtractImagesBasedOnCustomXml(string docPath)
    {
        var result = new Dictionary<string, string>();
        Document doc = new Document(docPath);

        // Iterate over all custom XML parts.
        foreach (CustomXmlPart customXmlPart in doc.CustomXmlParts)
        {
            // Parse the XML content.
            string xmlString = Encoding.UTF8.GetString(customXmlPart.Data);
            XDocument xDoc = XDocument.Parse(xmlString);
            XNamespace ns = "http://example.com/resources";

            var resources = xDoc.Root?.Elements(ns + "Resource");
            if (resources == null) continue;

            foreach (var resource in resources)
            {
                string imageName = resource.Element(ns + "ImageName")?.Value;
                string externalId = resource.Element(ns + "ExternalId")?.Value;

                if (string.IsNullOrEmpty(imageName) || string.IsNullOrEmpty(externalId))
                    continue;

                // Find the shape with the matching name.
                Shape shape = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .FirstOrDefault(s => s.Name == imageName);

                if (shape != null && shape.HasImage)
                {
                    string outputFileName = $"{externalId}.png";
                    shape.ImageData.Save(outputFileName);
                    if (!File.Exists(outputFileName))
                        throw new InvalidOperationException($"Failed to save extracted image to '{outputFileName}'.");

                    result[externalId] = outputFileName;
                }
            }
        }

        return result;
    }
}
