using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Add a custom XML part that will serve as the data source for the content controls.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
                                <Customer>
                                    <Name>John Doe</Name>
                                    <Age>30</Age>
                                </Customer>
                              </root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // 3. Insert a plain‑text content control mapped to the <Name> element.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Name[1]", string.Empty);
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(nameSdt);

        // 4. Insert another plain‑text content control mapped to the <Age> element.
        StructuredDocumentTag ageSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerAge",
            Tag = "customer-age"
        };
        ageSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Age[1]", string.Empty);
        para.AppendChild(new Run(doc, " "));
        para.AppendChild(ageSdt);

        // 5. Save the document so that the mappings are persisted.
        string docPath = "MappedContentControls.docx";
        doc.Save(docPath);

        // 6. Collect XML mapping information from all content controls.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>()
                          .Where(s => s.XmlMapping.IsMapped)
                          .ToList();

        // Extract distinct element names from the XPath expressions.
        HashSet<string> elementNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sdt in sdtNodes)
        {
            // Example XPath: /root[1]/Customer[1]/Name[1]
            string[] parts = sdt.XmlMapping.XPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                // Remove any index suffix like [1].
                string name = part.Split('[')[0];
                if (!string.IsNullOrEmpty(name))
                    elementNames.Add(name);
            }
        }

        // 7. Build a simple XSD schema that contains the collected element names.
        StringBuilder xsdBuilder = new StringBuilder();
        xsdBuilder.AppendLine(@"<?xml version=""1.0"" encoding=""utf-8""?>");
        xsdBuilder.AppendLine(@"<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">");
        xsdBuilder.AppendLine(@"  <xs:element name=""root"">");
        xsdBuilder.AppendLine(@"    <xs:complexType>");
        xsdBuilder.AppendLine(@"      <xs:sequence>");

        // Add child elements (excluding the root itself).
        foreach (string name in elementNames.Where(n => !n.Equals("root", StringComparison.OrdinalIgnoreCase)))
        {
            xsdBuilder.AppendLine($@"        <xs:element name=""{name}"" type=""xs:string"" minOccurs=""0""/>");
        }

        // Close the XSD tags.
        xsdBuilder.AppendLine(@"      </xs:sequence>");
        xsdBuilder.AppendLine(@"    </xs:complexType>");
        xsdBuilder.AppendLine(@"  </xs:element>");
        xsdBuilder.AppendLine(@"</xs:schema>");

        // 8. Save the generated XSD to a file.
        string xsdPath = "ContentControlsSchema.xsd";
        File.WriteAllText(xsdPath, xsdBuilder.ToString());

        // Optional: inform the user (console output is harmless in a console app).
        Console.WriteLine($"Document saved to: {docPath}");
        Console.WriteLine($"XSD schema saved to: {xsdPath}");
    }
}
