using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;          // For CustomXmlPart
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Prepare deterministic working folder and file names.
        // -------------------------------------------------
        const string workDir = "Work";
        const string imageFile = "sample.png";
        const string docFile = "sample.docx";

        Directory.CreateDirectory(workDir);
        string imagePath = Path.Combine(workDir, imageFile);
        string docPath = Path.Combine(workDir, docFile);

        // -------------------------------------------------
        // 2. Create a deterministic sample image (100x100 white PNG).
        // -------------------------------------------------
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
            }

            bitmap.Save(imagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 3. Build a DOCX that contains the image and a custom XML part
        //    mapping the shape's Name to an external resource identifier.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image and keep a reference to the created shape.
        Shape imageShape = builder.InsertImage(imagePath);

        // Assign a deterministic name to the shape (used as the key in the mapping).
        const string shapeName = "MyImageShape";
        imageShape.Name = shapeName;

        // Create custom XML that maps the shape Name to an external resource Id.
        const string schemaId = "urn:my-schema";
        XNamespace ns = schemaId;
        XDocument customXml = new XDocument(
            new XElement(ns + "ImageMap",
                new XElement(ns + "Image",
                    new XAttribute("shapeName", shapeName),
                    new XAttribute("externalId", "ext123"))));

        // Add the custom XML part to the document.
        // The Add method expects an Id (any unique string) and the XML content.
        string xmlPartId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(xmlPartId, customXml.ToString());

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 4. Load the document and read the custom XML mapping.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        // Build a dictionary: shape Name -> external resource Id.
        Dictionary<string, string> shapeNameToExternalId = new Dictionary<string, string>();

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // Load the XML content from the byte array.
            XDocument partXml = XDocument.Parse(Encoding.UTF8.GetString(part.Data));
            // The namespace used in the custom XML.
            XNamespace partNs = schemaId;

            // Find all Image elements.
            foreach (XElement imgElem in partXml.Descendants(partNs + "Image"))
            {
                XAttribute nameAttr = imgElem.Attribute("shapeName");
                XAttribute extAttr = imgElem.Attribute("externalId");
                if (nameAttr != null && extAttr != null)
                {
                    shapeNameToExternalId[nameAttr.Value] = extAttr.Value;
                }
            }
        }

        // -------------------------------------------------
        // 5. Extract images from shapes and save them using the external IDs.
        // -------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Determine external ID if a mapping exists.
            if (!shapeNameToExternalId.TryGetValue(shape.Name, out string externalId))
                continue; // No mapping – skip.

            // Determine file extension based on image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFileName = $"Image_{externalId}{extension}";
            string outPath = Path.Combine(workDir, outFileName);

            // Save the image.
            shape.ImageData.Save(outPath);
            extractedCount++;
        }

        // -------------------------------------------------
        // 6. Validation – ensure at least one image was extracted.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted based on the custom XML mapping.");
    }
}
