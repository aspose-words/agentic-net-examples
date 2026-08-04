using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlXmlBindingExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part that contains some sample data.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><name>John Doe</name></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Prepare a plain‑text content control (SDT) that will be bound to XML data.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PersonName",
            Tag = "person-name"
        };

        // Attempt to map the content control to an existing XML node.
        bool nameMapped = sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/name[1]", string.Empty);

        // Attempt to map the same content control to a missing XML node.
        // This will fail, and we will handle the situation gracefully.
        bool addressMapped = sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/address[1]", string.Empty);

        // If the mapping to the address node failed, replace the content with a placeholder message.
        if (!addressMapped)
        {
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "[Address not available]"));
        }

        // Insert the content control into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer Information:");
        builder.InsertNode(sdt);
        builder.Writeln(); // Add a line break after the control.

        // Save the resulting document.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputDocPath);

        // Prepare a simple status object to demonstrate JSON serialization.
        var mappingStatus = new
        {
            NameMappingSuccessful = nameMapped,
            AddressMappingSuccessful = addressMapped,
            OutputDocument = outputDocPath
        };

        // Serialize the status to JSON and write it to a file.
        string json = JsonConvert.SerializeObject(mappingStatus, Formatting.Indented);
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "mapping_status.json");
        File.WriteAllText(jsonPath, json, Encoding.UTF8);
    }
}
