using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -----------------------------------------------------------------
        // 1. Create a custom XML part that will hold the data for the SDTs.
        // -----------------------------------------------------------------
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent =
            "<root>" +
                "<person>" +
                    "<firstName>John</firstName>" +
                    "<lastName>Doe</lastName>" +
                "</person>" +
                "<address>" +
                    "<city>Seattle</city>" +
                    "<country>USA</country>" +
                "</address>" +
            "</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 2. Insert a few content controls and map them to nodes in the XML part.
        // -----------------------------------------------------------------
        // Helper to create a plain‑text SDT, map it and add it to the document body.
        void AddMappedContentControl(string title, string xpath)
        {
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = title,
                Tag = title.Replace(" ", "_")
            };
            // Map the SDT to the specified XPath within the custom XML part.
            sdt.XmlMapping.SetMapping(xmlPart, xpath, string.Empty);
            // Insert the SDT into the first paragraph.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(sdt);
        }

        // Ensure the document has at least one paragraph.
        if (doc.FirstSection.Body.FirstParagraph == null)
            doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        AddMappedContentControl("First Name", "/root[1]/person[1]/firstName[1]");
        AddMappedContentControl("Last Name", "/root[1]/person[1]/lastName[1]");
        AddMappedContentControl("City", "/root[1]/address[1]/city[1]");
        AddMappedContentControl("Country", "/root[1]/address[1]/country[1]");

        // Save the sample document (optional, just to visualize the result).
        doc.Save("SampleDocument.docx");

        // -----------------------------------------------------------------
        // 3. Gather XML mapping information from all content controls.
        // -----------------------------------------------------------------
        List<StructuredDocumentTag> sdtList = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                                .OfType<StructuredDocumentTag>()
                                                .ToList();

        // Collect distinct element names used in the mappings.
        HashSet<string> elementNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (StructuredDocumentTag sdt in sdtList)
        {
            if (sdt.XmlMapping.IsMapped)
            {
                // Example XPath: /root[1]/person[1]/firstName[1]
                // Split the path and take the element names (ignore indexes).
                string[] parts = sdt.XmlMapping.XPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string part in parts)
                {
                    // Remove any index like [1].
                    string name = part.Split('[')[0];
                    if (!string.IsNullOrWhiteSpace(name) && name != "root")
                        elementNames.Add(name);
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Build a simple XSD schema that defines the collected elements.
        // -----------------------------------------------------------------
        // This XSD is minimal and demonstrates the element structure.
        // For a real‑world scenario you would generate a complete schema
        // based on the full XML hierarchy and data types.
        string xsdTemplate = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""root"">
    <xs:complexType>
      <xs:sequence>
{0}
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Build the inner sequence with placeholder complex types for each element.
        List<string> elementDefinitions = new List<string>();
        foreach (string elem in elementNames)
        {
            string definition =
$@"        <xs:element name=""{elem}"">
          <xs:complexType>
            <xs:simpleContent>
              <xs:extension base=""xs:string""/>
            </xs:simpleContent>
          </xs:complexType>
        </xs:element>";
            elementDefinitions.Add(definition);
        }

        string xsdContent = string.Format(xsdTemplate, string.Join(Environment.NewLine, elementDefinitions));

        // -----------------------------------------------------------------
        // 5. Write the XSD to an external file.
        // -----------------------------------------------------------------
        string xsdPath = "ContentControlsMapping.xsd";
        File.WriteAllText(xsdPath, xsdContent);

        // Informative console output (no user interaction required).
        Console.WriteLine($"Generated XSD schema with {elementNames.Count} element definitions.");
        Console.WriteLine($"Schema saved to: {Path.GetFullPath(xsdPath)}");
    }
}
