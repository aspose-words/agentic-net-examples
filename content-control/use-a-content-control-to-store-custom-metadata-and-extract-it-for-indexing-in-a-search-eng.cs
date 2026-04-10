using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
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

        // Create a custom XML part that will hold metadata.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<meta><author>John Doe</author><department>Engineering</department></meta>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a block‑level plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "AuthorInfo",
            Tag = "AuthorMeta",
            IsShowingPlaceholderText = false
        };

        // Map the content control to the <author> element of the custom XML part.
        sdt.XmlMapping.SetMapping(xmlPart, "/meta[1]/author[1]", string.Empty);

        // Add the content control to the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Inside the content control, add a paragraph with a run.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "John Doe");
        para.AppendChild(run);
        sdt.AppendChild(para);

        // Save the document.
        string docPath = Path.Combine(Environment.CurrentDirectory, "SampleWithMetadata.docx");
        doc.Save(docPath);

        // Extract metadata from all content controls in the document.
        List<MetadataItem> items = new List<MetadataItem>();
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag tag in sdtNodes)
        {
            string title = tag.Title ?? string.Empty;
            string tagValue = tag.Tag ?? string.Empty;
            string displayText = tag.GetText().Trim();

            string? mappedXmlValue = null;
            if (tag.XmlMapping.IsMapped)
            {
                // Load the XML part data.
                string partXml = Encoding.UTF8.GetString(tag.XmlMapping.CustomXmlPart.Data);
                try
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(partXml);
                    XmlNode? node = xmlDoc.SelectSingleNode(tag.XmlMapping.XPath);
                    if (node != null)
                        mappedXmlValue = node.InnerText;
                }
                catch
                {
                    // Ignore XML parsing errors.
                }
            }

            items.Add(new MetadataItem
            {
                Title = title,
                Tag = tagValue,
                DisplayText = displayText,
                MappedXmlValue = mappedXmlValue
            });
        }

        // Serialize the extracted metadata to JSON.
        string json = JsonConvert.SerializeObject(items, Newtonsoft.Json.Formatting.Indented);
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "metadata.json");
        File.WriteAllText(jsonPath, json);

        // Output the JSON to the console.
        Console.WriteLine("Extracted metadata:");
        Console.WriteLine(json);
    }

    // Simple DTO for JSON serialization.
    private class MetadataItem
    {
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
        public string? MappedXmlValue { get; set; }
    }
}
