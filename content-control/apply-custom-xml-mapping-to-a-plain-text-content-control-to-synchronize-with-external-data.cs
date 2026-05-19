using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define custom XML data.
        string xml = "<root><Customer><Name>John Doe</Name><Email>john@example.com</Email></Customer></root>";

        // Add the custom XML part to the document using a GUID as the part ID.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xml);

        // Create a plain‑text inline content control for the customer's name.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        // Map the control to the <Name> element in the custom XML part.
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Name[1]", string.Empty);

        // Create a plain‑text inline content control for the customer's email.
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerEmail",
            Tag = "customer-email"
        };
        // Map the control to the <Email> element in the custom XML part.
        emailSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/Customer[1]/Email[1]", string.Empty);

        // Insert the controls into the first paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(new Run(doc, "Name: "));
        para.AppendChild(nameSdt);
        para.AppendChild(new Run(doc, "\nEmail: "));
        para.AppendChild(emailSdt);

        // Save the initial document.
        doc.Save("MappedContentControl.docx");

        // Load the document again to demonstrate updating the XML part.
        Document loadedDoc = new Document("MappedContentControl.docx");
        // Retrieve the same custom XML part by its ID.
        CustomXmlPart loadedPart = loadedDoc.CustomXmlParts.GetById(partId);
        // Update the XML content.
        string updatedXml = "<root><Customer><Name>Jane Smith</Name><Email>jane@example.com</Email></Customer></root>";
        loadedPart.Data = Encoding.UTF8.GetBytes(updatedXml);

        // Save the document after the XML update; the content controls reflect the new values.
        loadedDoc.Save("MappedContentControl_Updated.docx");
    }
}
