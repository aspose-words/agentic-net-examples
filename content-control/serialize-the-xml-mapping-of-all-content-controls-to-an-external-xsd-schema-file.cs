using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -----------------------------------------------------------------
        // 1. Create a custom XML part that will hold the data for content controls.
        // -----------------------------------------------------------------
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent =
            "<root>" +
                "<person>" +
                    "<firstName>John</firstName>" +
                    "<lastName>Doe</lastName>" +
                "</person>" +
                "<person>" +
                    "<firstName>Jane</firstName>" +
                    "<lastName>Smith</lastName>" +
                "</person>" +
            "</root>";

        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 2. Add an XSD schema that describes the structure of the XML part.
        // -----------------------------------------------------------------
        string xsdSchema =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
            <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
              <xs:element name=""root"">
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""person"" maxOccurs=""unbounded"">
                      <xs:complexType>
                        <xs:sequence>
                          <xs:element name=""firstName"" type=""xs:string""/>
                          <xs:element name=""lastName"" type=""xs:string""/>
                        </xs:sequence>
                      </xs:complexType>
                    </xs:element>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:schema>";

        // The Schemas collection stores the XSD strings.
        xmlPart.Schemas.Add(xsdSchema);

        // -----------------------------------------------------------------
        // 3. Insert plain‑text content controls and map them to XML nodes.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First person – first name.
        StructuredDocumentTag firstNameTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        firstNameTag.Title = "FirstName";
        firstNameTag.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/firstName[1]", string.Empty);

        builder.Writeln("First Name:");
        builder.InsertParagraph();               // Ensure we are inside a paragraph.
        builder.InsertNode(firstNameTag);        // Inline SDT can be inserted into a paragraph.

        // First person – last name.
        StructuredDocumentTag lastNameTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        lastNameTag.Title = "LastName";
        lastNameTag.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/lastName[1]", string.Empty);

        builder.Writeln("Last Name:");
        builder.InsertParagraph();
        builder.InsertNode(lastNameTag);

        // -----------------------------------------------------------------
        // 4. Serialize the XSD schema(s) associated with the custom XML part to a file.
        // -----------------------------------------------------------------
        string schemaOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "MappingSchema.xsd");
        using (StreamWriter writer = new StreamWriter(schemaOutputPath, false, System.Text.Encoding.UTF8))
        {
            foreach (string schema in xmlPart.Schemas)
            {
                writer.WriteLine(schema);
            }
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        string docOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "MappedDocument.docx");
        doc.Save(docOutputPath);
    }
}
