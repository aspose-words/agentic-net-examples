using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG).
        // -------------------------------------------------
        string workDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(workDir, "sample.png");
        string docPath = Path.Combine(workDir, "sample.docx");

        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }

            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Build a DOCX that contains the image and a
        //    custom XML part mapping the image name to an
        //    external resource ID.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image and give the shape a deterministic name.
        Shape imgShape = builder.InsertImage(imagePath);
        imgShape.Name = "Image1";

        // Create a custom XML part that maps the shape name to an external ID.
        string customXml = @"
            <Mappings xmlns='http://example.com/mappings'>
                <Mapping ImageName='Image1' ExternalId='Res123' />
            </Mappings>";

        // Add the custom XML part using the overload that accepts an ID and XML string.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, customXml);

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and parse the custom XML.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var nameToExternalId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // Convert the part's byte data to a string.
            string xmlContent = Encoding.UTF8.GetString(part.Data);
            XDocument xDoc = XDocument.Parse(xmlContent);
            XNamespace ns = "http://example.com/mappings";

            foreach (XElement mapping in xDoc.Descendants(ns + "Mapping"))
            {
                string imageName = (string)mapping.Attribute("ImageName");
                string externalId = (string)mapping.Attribute("ExternalId");
                if (!string.IsNullOrEmpty(imageName) && !string.IsNullOrEmpty(externalId))
                {
                    nameToExternalId[imageName] = externalId;
                }
            }
        }

        // -------------------------------------------------
        // 4. Extract images from shapes and save them using
        //    the external resource IDs from the custom XML.
        // -------------------------------------------------
        int extractedCount = 0;
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Use the shape's Name property for mapping.
            string key = shape.Name;
            if (string.IsNullOrEmpty(key) || !nameToExternalId.TryGetValue(key, out string externalId))
                continue; // No mapping found for this shape.

            // Determine file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFile = Path.Combine(workDir, $"{externalId}{extension}");

            // Save the image.
            shape.ImageData.Save(outFile);
            extractedCount++;
        }

        // -------------------------------------------------
        // 5. Validation – ensure at least one image was saved.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted based on the custom XML mapping.");
    }
}
