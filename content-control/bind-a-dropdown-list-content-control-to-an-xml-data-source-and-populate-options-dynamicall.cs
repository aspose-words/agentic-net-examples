using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define XML that holds the dropdown options.
        string xml = @"<options>
                         <option value=""A"">Option A</option>
                         <option value=""B"">Option B</option>
                         <option value=""C"">Option C</option>
                       </options>";

        // Add the XML as a custom XML part to the document.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xml);

        // Parse the XML to retrieve the option values.
        XDocument xDoc = XDocument.Parse(xml);
        // Create a dropdown list content control (inline level).
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "SampleDropdown",
            Tag = "sample-dropdown"
        };

        // Populate the dropdown list items from the XML.
        foreach (XElement option in xDoc.Root.Elements("option"))
        {
            string displayText = option.Value;
            string value = option.Attribute("value")?.Value ?? displayText;
            dropdown.ListItems.Add(new SdtListItem(displayText, value));
        }

        // Optionally set the initially selected value.
        if (dropdown.ListItems.Count > 0)
            dropdown.ListItems.SelectedValue = dropdown.ListItems[0];

        // Insert the content control into the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(dropdown);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DropdownMapped.docx");
        doc.Save(outputPath);
    }
}
