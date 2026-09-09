using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (a blank one exists by default).
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Define custom XML data that will be mapped to content controls.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xml = @"<root>
  <customer>
    <name>John Doe</name>
    <email>john.doe@example.com</email>
  </customer>
</root>";

        // Add the custom XML part to the document.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xml);

        // Create a plain‑text content control for the customer's name.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        // Map the control to the <name> element.
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/name[1]", string.Empty);
        para.AppendChild(nameSdt);

        // Add a space between the two controls.
        para.AppendChild(new Run(doc, " "));

        // Create a plain‑text content control for the customer's email.
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerEmail",
            Tag = "customer-email"
        };
        // Map the control to the <email> element.
        emailSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/email[1]", string.Empty);
        para.AppendChild(emailSdt);

        // Save the document.
        const string outputPath = "CustomXmlMapped.docx";
        doc.Save(outputPath);

        // Load the saved document and print the mapped values of the content controls.
        Document loaded = new Document(outputPath);
        NodeCollection sdtNodes = loaded.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            Console.WriteLine($"{sdt.Title}: {sdt.GetText().Trim()}");
        }
    }
}
