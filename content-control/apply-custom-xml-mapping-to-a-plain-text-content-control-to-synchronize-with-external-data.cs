using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Customer Information:");

        // Insert the first plain‑text content control for the customer name.
        builder.Write("Name: ");
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        builder.InsertNode(nameSdt);
        builder.Writeln(); // Move to the next line.

        // Insert the second plain‑text content control for the order ID.
        builder.Write("Order ID: ");
        StructuredDocumentTag orderSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "OrderId",
            Tag = "order-id"
        };
        builder.InsertNode(orderSdt);
        builder.Writeln();

        // Create a custom XML part that holds the external data.
        string xml = @"
<root>
    <customer>
        <name>Contoso Ltd.</name>
        <orderId>12345</orderId>
    </customer>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xml);

        // Map each content control to the corresponding XML node.
        nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/name[1]", string.Empty);
        orderSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/orderId[1]", string.Empty);

        // Save the resulting document.
        doc.Save("MappedContentControl.docx");
    }
}
