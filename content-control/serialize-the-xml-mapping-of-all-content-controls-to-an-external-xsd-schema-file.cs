using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part that will be the data source for content controls.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
    <Customer>
        <Name>John Doe</Name>
        <Address>123 Main St</Address>
    </Customer>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a plain‑text content control mapped to the <Name> element.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Name[1]", string.Empty);
        doc.FirstSection.Body.FirstParagraph.AppendChild(nameSdt);

        // Insert a plain‑text content control mapped to the <Address> element.
        StructuredDocumentTag addressSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerAddress",
            Tag = "customer-address"
        };
        addressSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Address[1]", string.Empty);
        doc.FirstSection.Body.FirstParagraph.AppendChild(addressSdt);

        // Save the sample document.
        const string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Collect XML mapping information from all content controls in the document.
        var mappings = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                         .OfType<StructuredDocumentTag>()
                         .Where(sdt => sdt.XmlMapping.IsMapped)
                         .Select(sdt => new
                         {
                             Title = sdt.Title,
                             Tag = sdt.Tag,
                             XPath = sdt.XmlMapping.XPath,
                             PartId = sdt.XmlMapping.CustomXmlPart?.Id ?? string.Empty
                         })
                         .ToList();

        // Build a simple XSD schema that defines each mapped element as a string.
        var elementNames = new HashSet<string>();
        foreach (var map in mappings)
        {
            // Extract the last element name from the XPath (e.g., "/root[1]/Customer[1]/Name[1]" -> "Name").
            string[] parts = map.XPath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0)
            {
                string lastPart = parts[^1];
                int bracketIndex = lastPart.IndexOf('[');
                string elementName = bracketIndex > 0 ? lastPart.Substring(0, bracketIndex) : lastPart;
                elementNames.Add(elementName);
            }
        }

        var sb = new StringBuilder();
        sb.AppendLine(@"<?xml version=""1.0"" encoding=""utf-8""?>");
        sb.AppendLine(@"<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">");
        sb.AppendLine(@"  <xs:element name=""root"">");
        sb.AppendLine(@"    <xs:complexType>");
        sb.AppendLine(@"      <xs:sequence>");

        foreach (string name in elementNames)
        {
            sb.AppendLine($@"        <xs:element name=""{name}"" type=""xs:string"" minOccurs=""0"" />");
        }

        sb.AppendLine(@"      </xs:sequence>");
        sb.AppendLine(@"    </xs:complexType>");
        sb.AppendLine(@"  </xs:element>");
        sb.AppendLine(@"</xs:schema>");

        // Write the generated XSD to a file.
        const string xsdPath = "ContentControlsSchema.xsd";
        File.WriteAllText(xsdPath, sb.ToString(), Encoding.UTF8);
    }
}
