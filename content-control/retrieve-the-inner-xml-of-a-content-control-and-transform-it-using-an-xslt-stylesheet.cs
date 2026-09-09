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

        // Add a custom XML part that will be mapped to a content control.
        string xmlPartContent = "<root><greeting>Hello, World!</greeting></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlPartContent);

        // Create an inline plain‑text content control and map it to the <greeting> element.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "GreetingControl",
            Tag = "greeting"
        };
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/greeting[1]", string.Empty);

        // Insert the content control into the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);

        // Save the document (optional, just to demonstrate the full lifecycle).
        doc.Save("ContentControl.docx");

        // Retrieve the XML that is mapped to the content control.
        string mappedXml = Encoding.UTF8.GetString(sdt.XmlMapping.CustomXmlPart.Data);

        // Define a simple XSLT that renames <greeting> to <message>.
        string xsltString = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">
  <xsl:output method=""xml"" indent=""yes""/>
  <xsl:template match=""/root"">
    <root>
      <message><xsl:value-of select=""greeting""/></message>
    </root>
  </xsl:template>
</xsl:stylesheet>";

        // Load the XSLT from the string.
        XslCompiledTransform xslt = new XslCompiledTransform();
        using (StringReader xsltReader = new StringReader(xsltString))
        using (XmlReader xsltXmlReader = XmlReader.Create(xsltReader))
        {
            xslt.Load(xsltXmlReader);
        }

        // Perform the transformation.
        string transformedXml;
        using (StringReader xmlReader = new StringReader(mappedXml))
        using (XmlReader inputReader = XmlReader.Create(xmlReader))
        using (StringWriter resultWriter = new StringWriter())
        using (XmlWriter outputWriter = XmlWriter.Create(resultWriter, xslt.OutputSettings))
        {
            xslt.Transform(inputReader, outputWriter);
            transformedXml = resultWriter.ToString();
        }

        // Output the transformed XML to the console.
        Console.WriteLine("Transformed XML:");
        Console.WriteLine(transformedXml);

        // Optionally, save the transformed XML to a file.
        File.WriteAllText("Transformed.xml", transformedXml, Encoding.UTF8);
    }
}
