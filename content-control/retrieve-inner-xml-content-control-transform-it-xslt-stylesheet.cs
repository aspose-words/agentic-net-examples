using System;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlXsltExample
{
    static void Main()
    {
        // Load the Word document that contains a content control (structured document tag).
        Document doc = new Document("Input.docx");

        // Find the first StructuredDocumentTag in the document.
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
        if (sdt == null)
        {
            Console.WriteLine("No content control found in the document.");
            return;
        }

        // Retrieve the inner XML of the content control.
        // WordOpenXML returns the full Flat OPC XML for the node; we can extract the inner part.
        string sdtXml = sdt.WordOpenXML;

        // Load the retrieved XML into an XmlDocument.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(sdtXml);

        // Prepare the XSLT transformation.
        XslCompiledTransform xslt = new XslCompiledTransform();
        xslt.Load("Transform.xsl"); // Path to the XSLT stylesheet.

        // Perform the transformation and capture the result as a string.
        using (StringWriter resultWriter = new StringWriter())
        using (XmlWriter xmlWriter = XmlWriter.Create(resultWriter, xslt.OutputSettings))
        {
            xslt.Transform(xmlDoc, xmlWriter);
            string transformedResult = resultWriter.ToString();

            // Output the transformed XML.
            Console.WriteLine("Transformed XML:");
            Console.WriteLine(transformedResult);
        }

        // Optionally, save the document after any modifications (none in this example).
        doc.Save("Output.docx");
    }
}
