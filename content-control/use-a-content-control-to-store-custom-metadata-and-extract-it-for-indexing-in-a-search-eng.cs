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
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading paragraph.
            builder.Writeln("Product catalog:");

            // Create a plain‑text content control that will hold custom metadata.
            StructuredDocumentTag productInfoSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "ProductInfo",
                Tag = "product-info"
            };

            // Create a custom XML part that stores the metadata.
            string xml = @"<metadata>
                               <product>
                                   <id>123</id>
                                   <category>Gadgets</category>
                               </product>
                           </metadata>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xml);

            // Map the content control to the <id> element of the custom XML.
            productInfoSdt.XmlMapping.SetMapping(xmlPart, "/metadata[1]/product[1]/id[1]", string.Empty);

            // Insert the content control into the current paragraph.
            builder.InsertNode(productInfoSdt);

            // Save the document that contains the content control and the custom XML part.
            const string docPath = "metadata.docx";
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // Load the document and extract the metadata for indexing.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(docPath);

            // Locate the content control by its Title.
            StructuredDocumentTag? foundSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .FirstOrDefault(s => s.Title == "ProductInfo");

            // Extract the value displayed by the content control (the product ID).
            string? productIdFromSdt = foundSdt?.GetText().Trim();

            // Retrieve the custom XML part that holds the full metadata.
            CustomXmlPart? metadataPart = loadedDoc.CustomXmlParts.FirstOrDefault();
            string? productCategory = null;
            string? productId = null;

            if (metadataPart != null)
            {
                string xmlContent = Encoding.UTF8.GetString(metadataPart.Data);
                XDocument xDoc = XDocument.Parse(xmlContent);
                XElement? productElement = xDoc.Root?.Element("product");
                productId = productElement?.Element("id")?.Value;
                productCategory = productElement?.Element("category")?.Value;
            }

            // Build an object that represents the searchable metadata.
            var searchableMetadata = new
            {
                Id = productId,
                Category = productCategory,
                ContentControlValue = productIdFromSdt
            };

            // Serialize the metadata to JSON (could be sent to a search engine indexer).
            string json = JsonConvert.SerializeObject(searchableMetadata, Formatting.Indented);
            const string jsonPath = "metadata.json";
            File.WriteAllText(jsonPath, json);

            // Output the JSON to the console for demonstration purposes.
            Console.WriteLine("Extracted metadata:");
            Console.WriteLine(json);
        }
    }
}
