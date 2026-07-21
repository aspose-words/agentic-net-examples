using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom XML part that will serve as the data source.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><customer><name>John Doe</name></customer></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -------------------------------------------------
        // Content control bound to an existing XML node.
        // -------------------------------------------------
        StructuredDocumentTag nameSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        nameSdt.Title = "CustomerName";
        nameSdt.Tag = "customer-name";

        // Attempt to map the control to the <name> element.
        bool nameMapped = nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/name[1]", string.Empty);
        if (!nameMapped || !nameSdt.XmlMapping.IsMapped)
        {
            // Mapping failed – provide a fallback placeholder.
            nameSdt.RemoveAllChildren();
            Paragraph placeholderPara = new Paragraph(doc);
            placeholderPara.AppendChild(new Run(doc, "[Name not found]"));
            nameSdt.AppendChild(placeholderPara);
        }

        // Move to a new paragraph for the next control.
        builder.Writeln();

        // -------------------------------------------------
        // Content control bound to a missing XML node.
        // -------------------------------------------------
        StructuredDocumentTag addressSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        addressSdt.Title = "CustomerAddress";
        addressSdt.Tag = "customer-address";

        // Attempt to map the control to a non‑existent <address> element.
        bool addressMapped = addressSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/customer[1]/address[1]", string.Empty);
        if (!addressMapped || !addressSdt.XmlMapping.IsMapped)
        {
            // Mapping failed – insert a clear placeholder indicating the missing data.
            addressSdt.RemoveAllChildren();
            Paragraph placeholderPara = new Paragraph(doc);
            placeholderPara.AppendChild(new Run(doc, "[Address not available]"));
            addressSdt.AppendChild(placeholderPara);
        }

        // Save the resulting document.
        const string outputDoc = "output.docx";
        doc.Save(outputDoc);

        // Export mapping information to JSON for verification.
        var mappingInfo = new[]
        {
            new { Title = nameSdt.Title, IsMapped = nameSdt.XmlMapping.IsMapped, XPath = nameSdt.XmlMapping.XPath },
            new { Title = addressSdt.Title, IsMapped = addressSdt.XmlMapping.IsMapped, XPath = addressSdt.XmlMapping.XPath }
        };
        string json = JsonConvert.SerializeObject(mappingInfo, Formatting.Indented);
        File.WriteAllText("mappingInfo.json", json);
    }
}
