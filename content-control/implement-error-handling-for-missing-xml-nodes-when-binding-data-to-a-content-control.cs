using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and ensure it has a minimum structure.
        Document doc = new Document();
        doc.EnsureMinimum();

        // Add a custom XML part that will be used for data binding.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
                                <name>John Doe</name>
                                <email>john.doe@example.com</email>
                              </root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Prepare a plain‑text content control (SDT) that we will bind to the XML.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "EmployeeAddress",
            Tag = "employee-address"
        };

        // The XPath we intend to bind to. This node does NOT exist in the XML above.
        const string xPath = "/root[1]/address[1]";

        // Verify that the target XML node exists before attempting to map.
        bool nodeExists = XmlNodeExists(xmlPart, xPath);

        if (nodeExists)
        {
            // The node exists – set the mapping.
            sdt.XmlMapping.SetMapping(xmlPart, xPath, string.Empty);
        }
        else
        {
            // The node is missing – handle the error gracefully.
            // Display a placeholder message inside the content control.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "[Address not available]"));
        }

        // Insert the content control into the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(sdt);

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Helper method that checks whether a given XPath selects a node in the provided CustomXmlPart.
    private static bool XmlNodeExists(CustomXmlPart part, string xPath)
    {
        if (part == null || part.Data == null || part.Data.Length == 0 || string.IsNullOrEmpty(xPath))
            return false;

        try
        {
            // Load the XML data from the custom part into an XmlDocument.
            XmlDocument xmlDoc = new XmlDocument();

            // Convert the byte[] data to a string and load it.
            string xmlString = Encoding.UTF8.GetString(part.Data);
            xmlDoc.LoadXml(xmlString);

            // Use XPath to locate the node.
            XmlNode node = xmlDoc.SelectSingleNode(xPath);
            return node != null;
        }
        catch
        {
            // If any exception occurs while parsing or querying, treat it as a missing node.
            return false;
        }
    }
}
