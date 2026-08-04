using System;
using System.IO;
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

        // Add a custom XML part with sample data.
        string xmlData = "<root><greeting>Hello</greeting><name>World</name></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlData);

        // Create an inline plain‑text content control and map it to the <greeting> element.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "GreetingControl",
            Tag = "greeting"
        };
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/greeting[1]", string.Empty);

        // Insert the content control into the first paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);

        // Save the document (optional, demonstrates that the control works).
        doc.Save("ContentControl.docx");

        // Retrieve the inner XML of the content control.
        string sdtXml = sdt.WordOpenXML;

        // XSLT that extracts the text inside the control.
        string xsltString = @"<?xml version='1.0' encoding='UTF-8'?>
<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'
                xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                exclude-result-prefixes='w'>
  <xsl:output method='text'/>
  <xsl:template match='/'>
    <xsl:value-of select='//w:t'/>
  </xsl:template>
</xsl:stylesheet>";

        // Load the XSLT.
        XslCompiledTransform xslt = new XslCompiledTransform();
        using (XmlReader xsltReader = XmlReader.Create(new StringReader(xsltString)))
        {
            xslt.Load(xsltReader);
        }

        // Transform the content control XML.
        string result;
        using (StringReader sdtReader = new StringReader(sdtXml))
        using (XmlReader xmlReader = XmlReader.Create(sdtReader))
        using (StringWriter writer = new StringWriter())
        {
            xslt.Transform(xmlReader, null, writer);
            result = writer.ToString();
        }

        // Output the transformed result.
        Console.WriteLine("Transformed content control text: " + result);
    }
}
