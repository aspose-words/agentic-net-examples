using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part with sample data.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><name>Contoso</name></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a plain text content control and map it to an existing XML node.
        StructuredDocumentTag existingNodeSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ExistingNode",
            Tag = "existing-node"
        };
        // Attempt to map to a valid XPath.
        bool mappingResult = existingNodeSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/name[1]", string.Empty);
        if (!mappingResult || !existingNodeSdt.XmlMapping.IsMapped)
        {
            Console.WriteLine("Failed to map existingNodeSdt to the XML node.");
        }

        // Insert the content control into the first paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(existingNodeSdt);

        // Create another plain text content control and attempt to map it to a missing XML node.
        StructuredDocumentTag missingNodeSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "MissingNode",
            Tag = "missing-node"
        };
        // This XPath does not exist in the XML part.
        bool missingMappingResult = missingNodeSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/missing[1]", string.Empty);

        // Prepare an error report object.
        var errorReport = new
        {
            ControlTitle = missingNodeSdt.Title,
            ControlTag = missingNodeSdt.Tag,
            XPath = "/root[1]/missing[1]",
            MappingSuccessful = missingMappingResult && missingNodeSdt.XmlMapping.IsMapped,
            Message = missingMappingResult && missingNodeSdt.XmlMapping.IsMapped
                ? "Mapping succeeded."
                : "Mapping failed: XML node not found."
        };

        // Serialize the error report to JSON and save it.
        string jsonReport = JsonConvert.SerializeObject(errorReport, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputDir, "errorReport.json"), jsonReport);

        // Insert the second content control into the document (even if mapping failed).
        para.AppendChild(missingNodeSdt);

        // Save the resulting document.
        string docPath = Path.Combine(outputDir, "MappedContentControls.docx");
        doc.Save(docPath);
    }
}
