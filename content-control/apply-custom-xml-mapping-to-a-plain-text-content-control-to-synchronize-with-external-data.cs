using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define custom XML data.
        string xml = "<root><CustomerName>Contoso</CustomerName></root>";

        // Add the custom XML part to the document.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xml);

        // Create a plain text content control (structured document tag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Map the content control to the XML node.
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/CustomerName[1]", string.Empty);

        // Insert the content control into the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);

        // Save the resulting document.
        doc.Save("CustomXmlMappedPlainText.docx");
    }
}
