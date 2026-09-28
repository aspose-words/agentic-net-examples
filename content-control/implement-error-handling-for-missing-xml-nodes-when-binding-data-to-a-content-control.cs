using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure there is at least one paragraph to host inline content controls.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        if (paragraph == null)
        {
            paragraph = new Paragraph(doc);
            doc.FirstSection.Body.AppendChild(paragraph);
        }

        // Define sample custom XML with only an existing node.
        string xmlContent = "<root><existing>Existing Value</existing></root>";

        // Add a custom XML part to the document (requires a unique ID).
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString(), xmlContent);

        // Load the XML into an XmlDocument for XPath queries.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);

        // Prepare a list of content control mappings to attempt.
        var mappings = new List<(string Title, string XPath)>
        {
            ("ExistingNodeControl", "/root[1]/existing[1]"),
            ("MissingNodeControl", "/root[1]/missing[1]")
        };

        // Collect binding results for JSON reporting.
        var bindingResults = new List<object>();

        foreach (var (title, xpath) in mappings)
        {
            try
            {
                // Attempt to locate the XML node using the provided XPath.
                XmlNode? xmlNode = xmlDoc.SelectSingleNode(xpath);
                if (xmlNode == null)
                {
                    // Node not found – record the error and continue.
                    bindingResults.Add(new
                    {
                        Title = title,
                        XPath = xpath,
                        Success = false,
                        Message = "XML node not found."
                    });
                    continue;
                }

                // Node exists – create an inline plain‑text content control.
                StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
                {
                    Title = title,
                    Tag = title
                };

                // Map the content control to the found XML node.
                sdt.XmlMapping.SetMapping(xmlPart, xpath, string.Empty);

                // Insert the content control into the paragraph.
                paragraph.AppendChild(sdt);

                // Record successful binding.
                bindingResults.Add(new
                {
                    Title = title,
                    XPath = xpath,
                    Success = true,
                    Message = "Binding succeeded."
                });
            }
            catch (Exception ex)
            {
                // Record any unexpected exceptions.
                bindingResults.Add(new
                {
                    Title = title,
                    XPath = xpath,
                    Success = false,
                    Message = $"Exception: {ex.Message}"
                });
            }
        }

        // Save the resulting document.
        const string outputDocPath = "BoundContentControls.docx";
        doc.Save(outputDocPath);

        // Serialize the binding results to a JSON file.
        string jsonOutput = JsonConvert.SerializeObject(bindingResults, Newtonsoft.Json.Formatting.Indented);
        File.WriteAllText("BindingResults.json", jsonOutput);
    }
}
