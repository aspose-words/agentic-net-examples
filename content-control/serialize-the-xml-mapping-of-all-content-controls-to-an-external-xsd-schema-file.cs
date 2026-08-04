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
        string xmlContent = @"<root>
    <person>
        <name>John Doe</name>
        <age>30</age>
    </person>
    <person>
        <name>Jane Smith</name>
        <age>28</age>
    </person>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 2. Define an XSD schema that describes the XML structure.
        //    The schema is stored externally; we will also write it to a file.
        // -----------------------------------------------------------------
        string xsdSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema' targetNamespace='http://example.com' xmlns='http://example.com' elementFormDefault='qualified'>
  <xs:element name='root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='person' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='name' type='xs:string'/>
              <xs:element name='age' type='xs:int'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Associate the schema URI with the custom XML part (required by Word).
        // The actual schema content will be saved separately.
        xmlPart.Schemas.Add("http://example.com");

        // -----------------------------------------------------------------
        // 3. Insert content controls (structured document tags) and map them to XML nodes.
        // -----------------------------------------------------------------
        // First content control – maps to the first person's name.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstPersonName",
            Tag = "first-person-name"
        };
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/name[1]", string.Empty);
        // Insert the control into the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(nameSdt);
        firstParagraph.AppendChild(new Run(doc, " ")); // space separator

        // Second content control – maps to the first person's age.
        StructuredDocumentTag ageSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstPersonAge",
            Tag = "first-person-age"
        };
        ageSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/age[1]", string.Empty);
        firstParagraph.AppendChild(ageSdt);
        firstParagraph.AppendChild(new Run(doc, "\n")); // new line

        // Third content control – maps to the second person's name.
        StructuredDocumentTag nameSdt2 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SecondPersonName",
            Tag = "second-person-name"
        };
        nameSdt2.XmlMapping.SetMapping(xmlPart, "/root[1]/person[2]/name[1]", string.Empty);
        firstParagraph.AppendChild(nameSdt2);
        firstParagraph.AppendChild(new Run(doc, " "));

        // Fourth content control – maps to the second person's age.
        StructuredDocumentTag ageSdt2 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SecondPersonAge",
            Tag = "second-person-age"
        };
        ageSdt2.XmlMapping.SetMapping(xmlPart, "/root[1]/person[2]/age[1]", string.Empty);
        firstParagraph.AppendChild(ageSdt2);

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        const string docPath = "MappedContentControls.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 5. Serialize the XSD schema to an external file.
        // -----------------------------------------------------------------
        const string xsdPath = "PersonSchema.xsd";
        File.WriteAllText(xsdPath, xsdSchema);

        // -----------------------------------------------------------------
        // 6. (Optional) Output a simple confirmation to the console.
        // -----------------------------------------------------------------
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"XSD schema saved to: {Path.GetFullPath(xsdPath)}");
    }
}
