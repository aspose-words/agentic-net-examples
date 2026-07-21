using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlMetadataExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for output files.
            const string docPath = "sample.docx";
            const string jsonPath = "metadata.json";

            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Add a custom XML part that holds product metadata.
            string xmlContent =
                "<metadata>" +
                "  <product>" +
                "    <name>Product A</name>" +
                "    <category>Electronics</category>" +
                "    <price>199.99</price>" +
                "  </product>" +
                "</metadata>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlContent);

            // 3. Insert a plain‑text content control that will display the product name.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            StructuredDocumentTag nameControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "ProductName",
                Tag = "product-name"
            };
            // Map the control to the <name> element of the custom XML part.
            nameControl.XmlMapping.SetMapping(xmlPart, "/metadata[1]/product[1]/name[1]", string.Empty);
            paragraph.AppendChild(nameControl);

            // 4. Save the document.
            doc.Save(docPath);

            // 5. Load the document back for extraction.
            Document loadedDoc = new Document(docPath);

            // 6. Locate the content control by its Title.
            StructuredDocumentTag? foundControl = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .FirstOrDefault(c => c.Title == "ProductName");

            if (foundControl == null)
                throw new InvalidOperationException("Content control not found.");

            // 7. Extract the displayed product name.
            string productName = foundControl.GetText().Trim();

            // 8. Retrieve the associated custom XML part.
            CustomXmlPart? associatedPart = foundControl.XmlMapping.CustomXmlPart;
            if (associatedPart == null)
                throw new InvalidOperationException("Custom XML part not found.");

            // 9. Parse the XML data.
            XDocument xDoc = XDocument.Parse(Encoding.UTF8.GetString(associatedPart.Data));

            // 10. Extract additional metadata (category and price) using LINQ to XML.
            XElement? productElement = xDoc.Root?.Element("product");
            if (productElement == null)
                throw new InvalidOperationException("Product element missing in XML.");

            string category = productElement.Element("category")?.Value ?? string.Empty;
            string priceText = productElement.Element("price")?.Value ?? "0";
            decimal price = decimal.TryParse(priceText, out decimal parsedPrice) ? parsedPrice : 0m;

            // 11. Prepare an object for JSON serialization.
            var metadata = new
            {
                Name = productName,
                Category = category,
                Price = price
            };

            // 12. Serialize to JSON and write to a file.
            string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);
            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // Optional: write a short confirmation to the console.
            Console.WriteLine("Metadata extracted and saved to " + jsonPath);
        }
    }
}
