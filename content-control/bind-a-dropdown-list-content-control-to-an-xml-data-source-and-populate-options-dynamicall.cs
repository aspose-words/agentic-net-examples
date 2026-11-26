using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Sample XML data source containing options and a selected value.
        string xmlData = @"
<root>
    <options>
        <option value='A'>Option A</option>
        <option value='B'>Option B</option>
        <option value='C'>Option C</option>
    </options>
    <selected>B</selected>
</root>";

        // Add the XML as a custom XML part to the document.
        // The first argument is a unique ID for the part.
        CustomXmlPart customXmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString(), xmlData);

        // Load the XML into an XmlDocument for easy traversal.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlData);

        // Create an inline dropdown list content control.
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "DynamicDropdown",
            Tag = "dynamic-dropdown"
        };

        // Populate the dropdown list items from the XML <option> elements.
        XmlNodeList? optionNodes = xmlDoc.SelectNodes("//option");
        if (optionNodes != null)
        {
            foreach (XmlNode optionNode in optionNodes)
            {
                string displayText = optionNode.InnerText ?? string.Empty;
                string value = optionNode.Attributes?["value"]?.Value ?? string.Empty;
                dropdown.ListItems.Add(new SdtListItem(displayText, value));
            }
        }

        // Map the content control's value to the <selected> element in the XML part.
        dropdown.XmlMapping.SetMapping(customXmlPart, "/root[1]/selected[1]", string.Empty);

        // Insert the dropdown into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(dropdown);

        // Save the resulting document.
        const string outputPath = "DropdownMapped.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
