using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define sample custom XML data.
        string customXml = @"<root>
    <Customer>
        <Name>Contoso</Name>
        <Address>123 Main St</Address>
    </Customer>
</root>";

        // Add the custom XML part to the document. Use a GUID as the part ID.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString(), customXml);

        // Insert a plain‑text content control mapped to the <Name> element.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Name[1]", "");
        nameSdt.RemoveAllChildren();
        nameSdt.AppendChild(new Run(doc, "Contoso"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(nameSdt);

        // Insert a plain‑text content control mapped to the <Address> element.
        StructuredDocumentTag addressSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerAddress",
            Tag = "customer-address"
        };
        addressSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Address[1]", "");
        addressSdt.RemoveAllChildren();
        addressSdt.AppendChild(new Run(doc, "123 Main St"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(addressSdt);

        // Save the document with mapped content controls.
        const string docPath = "MappedContentControls.docx";
        doc.Save(docPath);

        // Enumerate all content controls and collect their XML mapping information.
        var mappings = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                XPath = sdt.XmlMapping?.XPath
            })
            .Where(m => !string.IsNullOrEmpty(m.XPath))
            .ToList();

        // Build a simple XSD schema based on the collected mappings.
        // For this example we create a static schema that matches the sample XML.
        XNamespace xs = "http://www.w3.org/2001/XMLSchema";
        XElement schemaElement =
            new XElement(xs + "schema",
                new XAttribute("elementFormDefault", "qualified"),
                new XElement(xs + "element",
                    new XAttribute("name", "root"),
                    new XElement(xs + "complexType",
                        new XElement(xs + "sequence",
                            new XElement(xs + "element",
                                new XAttribute("name", "Customer"),
                                new XElement(xs + "complexType",
                                    new XElement(xs + "sequence",
                                        new XElement(xs + "element",
                                            new XAttribute("name", "Name"),
                                            new XAttribute("type", "xs:string")),
                                        new XElement(xs + "element",
                                            new XAttribute("name", "Address"),
                                            new XAttribute("type", "xs:string"))
                                    )
                                )
                            )
                        )
                    )
                )
            );

        // Create an XDocument with XML declaration.
        XDocument schemaDoc = new XDocument(
            new XDeclaration("1.0", "utf-8", "yes"),
            schemaElement
        );

        // Save the XSD schema to a file.
        const string xsdPath = "CustomSchema.xsd";
        File.WriteAllText(xsdPath, schemaDoc.ToString());

        // Output paths and mapping details for verification.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"XSD schema saved to: {Path.GetFullPath(xsdPath)}");
        Console.WriteLine("Content control mappings:");
        foreach (var m in mappings)
        {
            Console.WriteLine($"- Title: {m.Title}, Tag: {m.Tag}, XPath: {m.XPath}");
        }
    }
}
