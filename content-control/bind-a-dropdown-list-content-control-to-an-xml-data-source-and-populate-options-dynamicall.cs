using System;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // XML that defines the dropdown options.
        string xml = @"<options>
                         <option value='A'>Option A</option>
                         <option value='B'>Option B</option>
                         <option value='C'>Option C</option>
                       </options>";

        // Add the XML as a custom XML part (optional, shows how to embed XML in the document).
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xml);

        // Parse the XML to retrieve the option elements.
        XDocument xDoc = XDocument.Parse(xml);
        var optionElements = xDoc.Root?.Elements("option");

        // Create a drop‑down list content control (inline level).
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "SampleDropdown",
            Tag = "sample-dropdown"
        };

        // Populate the dropdown list items from the XML data.
        if (optionElements != null)
        {
            foreach (var opt in optionElements)
            {
                string displayText = opt.Value;
                string value = (string)opt.Attribute("value") ?? displayText;
                dropdown.ListItems.Add(new SdtListItem(displayText, value));
            }
        }

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(dropdown);

        // Save the resulting document.
        doc.Save("DropdownBound.docx");
    }
}
