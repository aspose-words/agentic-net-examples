using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlXmlTransform
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a custom XML part.
            Document doc = new Document();

            // Add a custom XML part containing simple data.
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlContent = @"<root><greeting>Hello, World!</greeting></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

            // Insert a block‑level plain‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
            {
                Title = "GreetingControl",
                Tag = "greeting"
            };
            // Map the content control to the <greeting> element in the custom XML part.
            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/greeting[1]", string.Empty);
            doc.FirstSection.Body.AppendChild(sdt);

            // Save the document to a file.
            const string docPath = "sample.docx";
            doc.Save(docPath);

            // Load the document back.
            Document loadedDoc = new Document(docPath);

            // Find the content control by its title.
            StructuredDocumentTag targetSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .FirstOrDefault(tag => tag.Title == "GreetingControl");

            if (targetSdt == null)
                throw new InvalidOperationException("Content control not found.");

            // Retrieve the inner XML of the content control.
            // WordOpenXML returns the full XML representation of the SDT node.
            string sdtXml = targetSdt.WordOpenXML;

            // Prepare a simple XSLT stylesheet that wraps the content in a <Transformed> element.
            const string xsltPath = "transform.xslt";
            string xsltContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <xsl:output method=""xml"" indent=""yes""/>
  <xsl:template match=""/"">
    <Transformed>
      <xsl:copy-of select=""*""/>
    </Transformed>
  </xsl:template>
</xsl:stylesheet>";
            File.WriteAllText(xsltPath, xsltContent);

            // Perform the XSLT transformation.
            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load(xsltPath);

            using (StringReader sr = new StringReader(sdtXml))
            using (XmlReader xmlReader = XmlReader.Create(sr))
            using (StringWriter sw = new StringWriter())
            using (XmlWriter xmlWriter = XmlWriter.Create(sw, xslt.OutputSettings))
            {
                xslt.Transform(xmlReader, xmlWriter);
                string transformedResult = sw.ToString();

                // Output the transformed XML to the console.
                Console.WriteLine("Transformed XML:");
                Console.WriteLine(transformedResult);
            }

            // Clean up temporary files (optional).
            // File.Delete(docPath);
            // File.Delete(xsltPath);
        }
    }
}
