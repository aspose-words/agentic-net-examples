using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define custom XML data that will be used for mapping.
        string xml = "<root><Customer><Name>Contoso Ltd.</Name><Address>123 Main St.</Address></Customer></root>";

        // Add the custom XML part to the document. Use a GUID as the part identifier.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xml);

        // Create an inline plain‑text content control (SDT).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };

        // Map the content control to the <Name> element inside the custom XML part.
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Name[1]", string.Empty);

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(sdt);

        // Save the resulting document.
        doc.Save("XmlMappedPlainTextSdt.docx");
    }
}
