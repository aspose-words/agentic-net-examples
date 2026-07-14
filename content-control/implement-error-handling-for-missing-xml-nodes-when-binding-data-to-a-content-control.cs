using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part that will be used for data binding.
        string xmlContent = "<root><name>John Doe</name></root>";
        string xmlPartId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 1. Content control bound to an existing XML node ("/root[1]/name[1]").
        // -----------------------------------------------------------------
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Name",
            Tag = "name"
        };

        // SetMapping returns true if the mapping succeeded.
        bool nameMapped = nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/name[1]", string.Empty);
        if (!nameMapped || !nameSdt.XmlMapping.IsMapped)
        {
            // This block will not be hit in this example because the node exists.
            nameSdt.RemoveAllChildren();
            nameSdt.AppendChild(new Run(doc, "[Name not found]"));
        }

        // Insert the content control into the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(nameSdt);

        // -----------------------------------------------------------------
        // 2. Content control bound to a missing XML node ("/root[1]/age[1]").
        // -----------------------------------------------------------------
        StructuredDocumentTag ageSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Age",
            Tag = "age"
        };

        try
        {
            // Attempt to map to a node that does not exist.
            bool ageMapped = ageSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/age[1]", string.Empty);

            // If mapping failed, throw an exception to be caught below.
            if (!ageMapped || !ageSdt.XmlMapping.IsMapped)
                throw new InvalidOperationException("The XML node for XPath '/root[1]/age[1]' was not found.");
        }
        catch (Exception ex)
        {
            // Provide a clear placeholder indicating the missing data.
            ageSdt.RemoveAllChildren();
            ageSdt.AppendChild(new Run(doc, $"[Missing data: {ex.Message}]"));
        }

        // Insert the second content control after the first one.
        firstParagraph.AppendChild(ageSdt);

        // Save the resulting document.
        doc.Save("output.docx");
    }
}
