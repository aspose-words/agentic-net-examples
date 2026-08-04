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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Add a plain‑text content control that holds a product name.
        // -----------------------------------------------------------------
        StructuredDocumentTag productNameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ProductName",
            Tag = "product-name"
        };
        productNameSdt.RemoveAllChildren();
        productNameSdt.AppendChild(new Run(doc, "Aspose.Words"));
        // Insert the control into the first paragraph.
        Paragraph firstPara = doc.FirstSection.Body.FirstParagraph;
        firstPara.AppendChild(productNameSdt);

        // -----------------------------------------------------------------
        // 2. Create a custom XML part that stores additional metadata.
        // -----------------------------------------------------------------
        string xmlContent = "<metadata><keywords>content control,metadata,search</keywords></metadata>";
        string xmlPartId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 3. Add a content control that is mapped to the XML part.
        // -----------------------------------------------------------------
        StructuredDocumentTag keywordsSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Keywords",
            Tag = "keywords"
        };
        // Map the control to the <keywords> element inside the custom XML part.
        keywordsSdt.XmlMapping.SetMapping(xmlPart, "/metadata[1]/keywords[1]", string.Empty);
        // Insert after the first control.
        firstPara.AppendChild(new Run(doc, " "));
        firstPara.AppendChild(keywordsSdt);

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        string docPath = "output.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 5. Extract metadata from all content controls for indexing.
        // -----------------------------------------------------------------
        List<MetadataItem> extracted = new List<MetadataItem>();

        // Enumerate all StructuredDocumentTag nodes in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
        {
            // Retrieve the displayed text of the control.
            string text = sdt.GetText().Trim();

            // If the control is mapped to XML, the text reflects the mapped value.
            extracted.Add(new MetadataItem
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = text
            });
        }

        // -----------------------------------------------------------------
        // 6. Serialize the extracted metadata to JSON for a search engine.
        // -----------------------------------------------------------------
        string json = JsonConvert.SerializeObject(extracted, Formatting.Indented);
        File.WriteAllText("metadata.json", json);
    }

    // Simple DTO for JSON output.
    private class MetadataItem
    {
        public string Title { get; set; }
        public string Tag { get; set; }
        public string Text { get; set; }
    }
}
