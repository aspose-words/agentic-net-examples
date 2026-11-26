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

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        sdt.Title = "SampleControl";
        sdt.Tag = "sample";

        // Add a paragraph with some text inside the content control.
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "Hello, World!"));
        sdt.AppendChild(innerParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the document in Flat OPC format to a memory stream.
        using (MemoryStream flatOpcStream = new MemoryStream())
        {
            doc.Save(flatOpcStream, SaveFormat.FlatOpc);
            flatOpcStream.Position = 0;

            // Load the Flat OPC XML.
            XmlDocument flatOpcXml = new XmlDocument();
            flatOpcXml.Load(flatOpcStream);

            // XSLT stylesheet that extracts the inner content of the SDT.
            string xsltString = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform""
                xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <xsl:output method=""xml"" indent=""yes""/>
  <!-- Match the SDT element and output its inner content -->
  <xsl:template match=""/w:sdt"">
    <TransformedContent>
      <xsl:apply-templates select=""w:sdtContent/*""/>
    </TransformedContent>
  </xsl:template>
  <!-- Copy all other elements and text nodes unchanged -->
  <xsl:template match=""*"">
    <xsl:copy>
      <xsl:apply-templates select=""@*|node()""/>
    </xsl:copy>
  </xsl:template>
  <xsl:template match=""@*|text()"">
    <xsl:copy/>
  </xsl:template>
</xsl:stylesheet>";

            // Load the XSLT from the string.
            XslCompiledTransform xslt = new XslCompiledTransform();
            using (StringReader xsltReader = new StringReader(xsltString))
            using (XmlReader xsltXmlReader = XmlReader.Create(xsltReader))
            {
                xslt.Load(xsltXmlReader);
            }

            // Perform the transformation on the Flat OPC XML.
            string transformedXml;
            using (StringWriter resultWriter = new StringWriter())
            using (XmlReader xmlReader = new XmlNodeReader(flatOpcXml))
            {
                xslt.Transform(xmlReader, null, resultWriter);
                transformedXml = resultWriter.ToString();
            }

            // Write the transformed XML to a file.
            File.WriteAllText("transformed.xml", transformedXml);
        }

        // Optionally save the original document for inspection.
        doc.Save("sample.docx");
    }
}
