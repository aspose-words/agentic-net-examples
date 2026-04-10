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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom XML part that will be mapped to a content control.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
    <item>
        <name>Apple</name>
        <price>1.20</price>
    </item>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a block‑level plain‑text content control and map it to the <item> element.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "ItemControl",
            Tag = "ItemTag"
        };
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/item[1]", string.Empty);
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the document (optional, demonstrates that the control exists in the file).
        doc.Save("ContentControl.docx");

        // Retrieve the inner XML of the mapped node using the custom XML part.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);
        XmlNode mappedNode = xmlDoc.SelectSingleNode("/root/item");
        string innerXml = mappedNode?.OuterXml ?? string.Empty;

        // Define a simple XSLT that wraps the original node in a <product> element.
        string xsltString = @"<?xml version='1.0'?>
<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
  <xsl:output method='xml' indent='yes'/>
  <xsl:template match='/'>
    <product>
      <xsl:copy-of select='*'/>
    </product>
  </xsl:template>
</xsl:stylesheet>";

        // Load the XSLT.
        XslCompiledTransform xslt = new XslCompiledTransform();
        using (StringReader sr = new StringReader(xsltString))
        using (XmlReader xr = XmlReader.Create(sr))
        {
            xslt.Load(xr);
        }

        // Transform the inner XML.
        string transformedXml;
        using (StringReader sr = new StringReader(innerXml))
        using (XmlReader xr = XmlReader.Create(sr))
        using (StringWriter sw = new StringWriter())
        {
            xslt.Transform(xr, null, sw);
            transformedXml = sw.ToString();
        }

        // Save the transformed XML to a file.
        File.WriteAllText("Transformed.xml", transformedXml, Encoding.UTF8);

        // Output the results to the console.
        Console.WriteLine("Inner XML of the content control:");
        Console.WriteLine(innerXml);
        Console.WriteLine();
        Console.WriteLine("Transformed XML:");
        Console.WriteLine(transformedXml);
    }
}
