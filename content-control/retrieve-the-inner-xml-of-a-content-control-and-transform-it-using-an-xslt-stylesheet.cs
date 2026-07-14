using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with a custom XML part.
        Document doc = new Document();

        // Sample XML data that will be mapped to the content control.
        string xmlContent = @"<root>
    <person>
        <firstName>John</firstName>
        <lastName>Doe</lastName>
        <age>30</age>
    </person>
</root>";

        // Add the XML part to the document.
        string xmlPartId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a plain‑text content control (SDT) into the first paragraph.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PersonInfo",
            Tag = "person-info"
        };
        // Map the SDT to the <person> element in the custom XML part.
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]", string.Empty);

        // Add the SDT to the document body.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Step 2: Load the document and locate the content control.
        Document loadedDoc = new Document(docPath);
        StructuredDocumentTag? targetSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(tag => tag.Title == "PersonInfo");

        if (targetSdt == null)
        {
            Console.WriteLine("Content control not found.");
            return;
        }

        // Step 3: Retrieve the inner XML of the mapped node.
        string innerXml = string.Empty;
        if (targetSdt.XmlMapping.IsMapped)
        {
            CustomXmlPart mappedPart = targetSdt.XmlMapping.CustomXmlPart;
            string partXml = Encoding.UTF8.GetString(mappedPart.Data);
            XmlDocument partDoc = new XmlDocument();
            partDoc.LoadXml(partXml);

            // Use the XPath stored in the mapping to locate the node.
            string xpath = targetSdt.XmlMapping.XPath;
            XmlNode? mappedNode = partDoc.SelectSingleNode(xpath);
            if (mappedNode != null)
                innerXml = mappedNode.InnerXml;
        }

        if (string.IsNullOrEmpty(innerXml))
        {
            Console.WriteLine("Failed to retrieve inner XML.");
            return;
        }

        // Step 4: Prepare a simple XSLT that converts the person data to plain text.
        const string xsltPath = "transform.xslt";
        string xsltContent = @"<?xml version='1.0' encoding='UTF-8'?>
<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
  <xsl:output method='text' encoding='UTF-8'/>
  <xsl:template match='/'>
    First Name: <xsl:value-of select='firstName'/>&#10;
    Last Name: <xsl:value-of select='lastName'/>&#10;
    Age: <xsl:value-of select='age'/>
  </xsl:template>
</xsl:stylesheet>";
        File.WriteAllText(xsltPath, xsltContent, Encoding.UTF8);

        // Step 5: Apply the XSLT transformation to the inner XML.
        XslCompiledTransform xslt = new XslCompiledTransform();
        xslt.Load(xsltPath);

        using (StringReader sr = new StringReader($"<person>{innerXml}</person>"))
        using (XmlReader xmlReader = XmlReader.Create(sr))
        using (StringWriter sw = new StringWriter())
        {
            xslt.Transform(xmlReader, null, sw);
            string result = sw.ToString();

            // Save the transformation result.
            const string resultPath = "result.txt";
            File.WriteAllText(resultPath, result, Encoding.UTF8);
            Console.WriteLine($"Transformation completed. Result saved to '{resultPath}'.");
        }

        // Clean up temporary XSLT file (optional).
        if (File.Exists(xsltPath))
            File.Delete(xsltPath);
    }
}
