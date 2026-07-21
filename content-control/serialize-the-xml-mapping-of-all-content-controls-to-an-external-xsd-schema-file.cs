using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a simple custom XML part that will be used for mapping.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
    <person>
        <firstName>John</firstName>
        <lastName>Doe</lastName>
    </person>
    <person>
        <firstName>Jane</firstName>
        <lastName>Smith</lastName>
    </person>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a plain‑text content control mapped to the first person's first name.
        StructuredDocumentTag firstNameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstName",
            Tag = "first-name-1"
        };
        firstNameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/firstName[1]", string.Empty);
        doc.FirstSection.Body.FirstParagraph.AppendChild(firstNameSdt);

        // Insert a plain‑text content control mapped to the first person's last name.
        StructuredDocumentTag lastNameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "LastName",
            Tag = "last-name-1"
        };
        lastNameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/lastName[1]", string.Empty);
        doc.FirstSection.Body.FirstParagraph.AppendChild(lastNameSdt);

        // Save the document (optional, just to have a physical file).
        doc.Save("MappedContentControls.docx");

        // Collect mapping information from all content controls in the document.
        List<MappingInfo> mappings = new List<MappingInfo>();
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (Node node in sdtNodes)
        {
            if (node is StructuredDocumentTag sdt && sdt.XmlMapping.IsMapped)
            {
                mappings.Add(new MappingInfo
                {
                    Title = sdt.Title,
                    Tag = sdt.Tag,
                    XPath = sdt.XmlMapping.XPath,
                    PrefixMappings = sdt.XmlMapping.PrefixMappings
                });
            }
        }

        // Serialize the mapping information to a simple XSD‑like file.
        // For demonstration we write the original XML part data as the schema source.
        string xsdPath = "ContentControlsMapping.xsd";
        File.WriteAllText(xsdPath, xmlPart.Data != null ? Encoding.UTF8.GetString(xmlPart.Data) : string.Empty);

        // Additionally, write a comment block with the collected mapping details.
        StringBuilder commentBuilder = new StringBuilder();
        commentBuilder.AppendLine("<!--");
        commentBuilder.AppendLine("Mapping information for content controls in the document:");
        foreach (MappingInfo info in mappings)
        {
            commentBuilder.AppendLine($"Title: {info.Title}, Tag: {info.Tag}, XPath: {info.XPath}, PrefixMappings: {info.PrefixMappings}");
        }
        commentBuilder.AppendLine("-->");
        File.AppendAllText(xsdPath, commentBuilder.ToString());

        // The program finishes without waiting for user input.
    }

    // Helper class to hold mapping details.
    private class MappingInfo
    {
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string XPath { get; set; } = string.Empty;
        public string PrefixMappings { get; set; } = string.Empty;
    }
}
