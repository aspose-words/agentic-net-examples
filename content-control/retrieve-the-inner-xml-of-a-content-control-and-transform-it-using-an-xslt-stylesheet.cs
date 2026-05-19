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

        // Insert a plain‑text content control (SDT) with some sample text.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = "sample-tag"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Hello, Aspose!"));

        // Add the content control to the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(sdt);

        // Save the document to the working directory.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document back (demonstrates the load rule).
        Document loadedDoc = new Document(docPath);

        // Find the content control by its title.
        StructuredDocumentTag foundSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(tag => tag.Title == "SampleControl");

        if (foundSdt == null)
            throw new InvalidOperationException("Content control not found.");

        // Retrieve the inner XML of the content control.
        string innerXml = foundSdt.WordOpenXML; // Full XML representation.

        // Prepare a simple XSLT that extracts the text node value.
        const string xsltString = @"
<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
  <xsl:output method='text' encoding='UTF-8'/>
  <xsl:template match='*'>
    <xsl:value-of select='.'/>
  </xsl:template>
</xsl:stylesheet>";

        // Load the XSLT from the string.
        XslCompiledTransform xslt = new XslCompiledTransform();
        using (XmlReader xsltReader = XmlReader.Create(new StringReader(xsltString)))
        {
            xslt.Load(xsltReader);
        }

        // Perform the transformation.
        string transformedResult;
        using (StringReader xmlReader = new StringReader(innerXml))
        using (XmlReader reader = XmlReader.Create(xmlReader))
        using (StringWriter writer = new StringWriter())
        {
            xslt.Transform(reader, null, writer);
            transformedResult = writer.ToString();
        }

        // Write the transformed result to a text file.
        const string outputPath = "transformed.txt";
        File.WriteAllText(outputPath, transformedResult, Encoding.UTF8);
    }
}
