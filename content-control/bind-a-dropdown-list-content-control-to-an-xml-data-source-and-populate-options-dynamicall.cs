using System;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define XML data source containing dropdown options and a selected value.
        string xml = @"<root>
  <options>
    <option value='A'>Option A</option>
    <option value='B'>Option B</option>
    <option value='C'>Option C</option>
  </options>
  <selected>B</selected>
</root>";

        // Add the XML as a custom XML part to the document (provide an ID for the part).
        CustomXmlPart customXmlPart = doc.CustomXmlParts.Add("myXmlPart", xml);

        // Get the first paragraph (creates one if the document is empty).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline dropdown list content control.
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline);
        dropdown.Title = "DynamicDropdown";
        dropdown.Tag = "dynamic-dropdown";

        // Load the XML into an XmlDocument for parsing.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xml);

        // Populate the dropdown list items from the <option> elements.
        XmlNodeList? optionNodes = xmlDoc.SelectNodes("//option");
        if (optionNodes != null)
        {
            foreach (XmlNode node in optionNodes)
            {
                string displayText = node.InnerText;
                string value = node.Attributes?["value"]?.Value ?? displayText;
                dropdown.ListItems.Add(new SdtListItem(displayText, value));
            }
        }

        // Bind the dropdown's selected value to the <selected> element in the XML part.
        dropdown.XmlMapping.SetMapping(customXmlPart, "/root[1]/selected[1]", string.Empty);

        // Insert the dropdown content control into the paragraph.
        paragraph.AppendChild(dropdown);

        // Save the resulting document.
        doc.Save("DropdownBound.docx");
    }
}
