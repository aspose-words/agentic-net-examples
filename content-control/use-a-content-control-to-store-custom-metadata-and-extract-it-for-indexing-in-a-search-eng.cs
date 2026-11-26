using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level rich‑text content control to hold product information.
        StructuredDocumentTag productInfoSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "ProductInfo",
            Tag = "product-info"
        };

        // Add a paragraph with the product name inside the content control.
        Paragraph productNamePara = new Paragraph(doc);
        productNamePara.AppendChild(new Run(doc, "Widget A"));
        productInfoSdt.AppendChild(productNamePara);

        // Insert the block‑level content control into the document body.
        doc.FirstSection.Body.AppendChild(productInfoSdt);

        // Create a custom XML part that stores additional metadata.
        string customXml = @"<metadata>
                                <category>Gadgets</category>
                                <price>19.99</price>
                             </metadata>";
        // Add the XML part with a generated ID.
        string xmlPartId = Guid.NewGuid().ToString();
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, customXml);

        // Create a plain‑text inline content control that maps to the <category> element.
        StructuredDocumentTag categorySdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Category",
            Tag = "product-category"
        };
        // Map the content control to the custom XML part.
        categorySdt.XmlMapping.SetMapping(xmlPart, "/metadata[1]/category[1]", string.Empty);

        // Insert the mapped content control after the product name.
        productNamePara.AppendChild(new Run(doc, " - Category: "));
        productNamePara.AppendChild(categorySdt);

        // Save the document with the content controls.
        string docPath = "ProductInfo.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Extraction phase: collect metadata from all content controls.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        List<object> extractedMetadata = new List<object>();

        // Enumerate all StructuredDocumentTag nodes in the document.
        foreach (StructuredDocumentTag sdt in loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                                    .OfType<StructuredDocumentTag>())
        {
            // Retrieve the displayed text of the content control.
            string text = sdt.GetText().Trim();

            // Build a simple metadata object.
            var meta = new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = text
            };
            extractedMetadata.Add(meta);
        }

        // Serialize the collected metadata to JSON for indexing.
        string json = JsonConvert.SerializeObject(extractedMetadata, Formatting.Indented);
        string jsonPath = "metadata-index.json";
        File.WriteAllText(jsonPath, json);
    }
}
