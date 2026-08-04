using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define XML that will serve as the data source for the dropdown list.
        // Each <item> element contains a display text and a value attribute.
        string xmlContent = @"
            <root>
                <item value='A'>Option A</item>
                <item value='B'>Option B</item>
                <item value='C'>Option C</item>
            </root>";

        // Add the XML as a custom XML part to the document.
        // The part ID can be any GUID string.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

        // Parse the XML to extract the items for the dropdown.
        XDocument xDoc = XDocument.Parse(xmlContent);
        var items = xDoc.Root?.Elements("item");

        // Create a dropdown list content control (inline level).
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "SampleDropdown",
            Tag = "sample-dropdown"
        };

        // Populate the dropdown list with items from the XML.
        if (items != null)
        {
            foreach (var elem in items)
            {
                string displayText = elem.Value;
                string value = elem.Attribute("value")?.Value ?? displayText;
                dropdown.ListItems.Add(new SdtListItem(displayText, value));
            }

            // Optionally set the selected value to the first item.
            if (dropdown.ListItems.Count > 0)
                dropdown.ListItems.SelectedValue = dropdown.ListItems[0];
        }

        // Insert the dropdown into the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(dropdown);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DropdownMapped.docx");
        doc.Save(outputPath);
    }
}
