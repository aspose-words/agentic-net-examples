using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlMetadataExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a paragraph to host inline content controls.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Plain text content control with custom metadata.
        StructuredDocumentTag productNameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ProductName",
            Tag = "product-name"
        };
        productNameSdt.RemoveAllChildren();
        productNameSdt.AppendChild(new Run(doc, "Widget"));
        para.AppendChild(productNameSdt);

        // Rich text (block) content control with custom metadata.
        StructuredDocumentTag descriptionSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Description",
            Tag = "product-description"
        };
        Paragraph descPara = new Paragraph(doc);
        descPara.AppendChild(new Run(doc, "A useful widget for everyday tasks."));
        descriptionSdt.AppendChild(descPara);
        doc.FirstSection.Body.AppendChild(descriptionSdt);

        // Drop‑down list content control with custom metadata.
        StructuredDocumentTag categorySdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Category",
            Tag = "product-category"
        };
        categorySdt.ListItems.Add(new SdtListItem("Tools", "Tools"));
        categorySdt.ListItems.Add(new SdtListItem("Gadgets", "Gadgets"));
        categorySdt.ListItems.Add(new SdtListItem("Accessories", "Accessories"));
        // Set default selected value by inserting the corresponding text.
        categorySdt.RemoveAllChildren();
        categorySdt.AppendChild(new Run(doc, "Tools"));
        para.AppendChild(categorySdt);

        // Create a custom XML part and map it to a content control.
        string customXml = "<product><id>12345</id></product>";
        // Add the XML part with an identifier.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add("productXml", customXml);
        StructuredDocumentTag productIdSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ProductId",
            Tag = "product-id"
        };
        // Map the content control to the <id> element.
        productIdSdt.XmlMapping.SetMapping(xmlPart, "/product[1]/id[1]", string.Empty);
        para.AppendChild(productIdSdt);

        // Save the document.
        const string docPath = "product.docx";
        doc.Save(docPath);

        // Extract metadata from all content controls for indexing.
        List<object> indexEntries = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .Cast<StructuredDocumentTag>()
            .Select(sdt =>
            {
                string? xpath = null;
                if (sdt.XmlMapping != null && sdt.XmlMapping.IsMapped)
                {
                    xpath = sdt.XmlMapping.XPath;
                }

                return new
                {
                    Title = sdt.Title,
                    Tag = sdt.Tag,
                    Type = sdt.SdtType.ToString(),
                    Text = sdt.GetText().Trim(),
                    XmlMappingXPath = xpath
                };
            })
            .Cast<object>()
            .ToList();

        // Serialize to JSON.
        string json = JsonConvert.SerializeObject(indexEntries, Formatting.Indented);
        const string jsonPath = "index.json";
        File.WriteAllText(jsonPath, json);

        // Output paths for verification (no interactive input required).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"Index JSON saved to: {Path.GetFullPath(jsonPath)}");
    }
}
