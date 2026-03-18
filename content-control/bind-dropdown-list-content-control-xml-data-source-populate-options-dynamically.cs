using System;
using System.IO;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;

class DropdownBindingExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // XML data that will be stored in a custom XML part.
        // The <option> elements contain the display text and the value for each dropdown item.
        string xmlContent =
            @"<root>
                <options>
                    <option display='Apple' value='A' />
                    <option display='Banana' value='B' />
                    <option display='Cherry' value='C' />
                </options>
              </root>";

        // Add the XML data as a custom XML part to the document.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart customXmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

        // Create a DropDownList structured document tag (content control).
        StructuredDocumentTag dropDownTag = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Block);

        // Map the content control to the <option> nodes inside the custom XML part.
        // The XPath selects all <option> elements under /root/options.
        dropDownTag.XmlMapping.SetMapping(customXmlPart, "/root[1]/options[1]/option", string.Empty);

        // Load the XML into an XmlDocument for easy traversal.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);
        XmlNodeList optionNodes = xmlDoc.SelectNodes("/root/options/option");

        // Populate the dropdown list items dynamically from the XML.
        foreach (XmlNode node in optionNodes)
        {
            // Retrieve the display text and value attributes.
            string display = node.Attributes["display"]?.Value ?? string.Empty;
            string value = node.Attributes["value"]?.Value ?? string.Empty;

            // Add a new list item to the content control.
            // Use the constructor that accepts both display text and value.
            dropDownTag.ListItems.Add(new SdtListItem(display, value));
        }

        // Optionally set the initially selected item (first one if any).
        if (dropDownTag.ListItems.Count > 0)
            dropDownTag.ListItems.SelectedValue = dropDownTag.ListItems[0];

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(dropDownTag);

        // Save the resulting document.
        doc.Save("DropdownBoundToXml.docx");
    }
}
