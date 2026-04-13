using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;          // For CustomXmlPart
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string imageFileName = "sample.png";
        const string docFileName = "input.docx";

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
            bitmap.Save(imageFileName, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX that contains the image and a custom XML part
        //    mapping the shape name to an external resource ID.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image; the returned Shape is automatically appended.
        Shape imgShape = builder.InsertImage(imageFileName);
        imgShape.Name = "ImageShape1"; // Deterministic name for mapping.

        // Simple custom XML that maps the shape name to a resource ID.
        string customXml = @"
            <images xmlns=""http://example.com/schema"">
                <image>
                    <shapeName>ImageShape1</shapeName>
                    <resourceId>res-001</resourceId>
                </image>
            </images>";

        // Add the custom XML part to the document.
        string xmlPartId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(xmlPartId, customXml);

        // Save the document.
        doc.Save(docFileName);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images according to the custom XML.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docFileName);

        if (loadedDoc.CustomXmlParts.Count == 0)
            throw new InvalidOperationException("No custom XML parts found.");

        // Retrieve the first (and only) custom XML part.
        CustomXmlPart xmlPart = loadedDoc.CustomXmlParts[0];
        string xmlContent = System.Text.Encoding.UTF8.GetString(xmlPart.Data);

        // Parse XML and build a mapping: shape name -> resource ID.
        var shapeToResource = new Dictionary<string, string>();
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);
        XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
        nsMgr.AddNamespace("ns", "http://example.com/schema");
        XmlNodeList imageNodes = xmlDoc.SelectNodes("//ns:image", nsMgr);
        foreach (XmlNode node in imageNodes)
        {
            string shapeName = node.SelectSingleNode("ns:shapeName", nsMgr)?.InnerText;
            string resourceId = node.SelectSingleNode("ns:resourceId", nsMgr)?.InnerText;
            if (!string.IsNullOrEmpty(shapeName) && !string.IsNullOrEmpty(resourceId))
                shapeToResource[shapeName] = resourceId;
        }

        // Iterate over all shapes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Find the resource ID using the shape name.
            if (!shapeToResource.TryGetValue(shape.Name, out string resourceId))
                continue; // No mapping for this shape.

            // Determine file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFileName = $"{resourceId}{extension}";

            // Save the image.
            shape.ImageData.Save(outFileName);
            extractedCount++;

            // Validate that the file was created.
            if (!File.Exists(outFileName))
                throw new InvalidOperationException($"Failed to save extracted image '{outFileName}'.");
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted based on the custom XML mapping.");

        // Optional cleanup (commented out).
        // File.Delete(imageFileName);
        // File.Delete(docFileName);
    }
}
