using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace DropdownXmlBindingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define XML that contains the list items and the initially selected value.
            string xmlContent = @"
                <root>
                    <options>
                        <option display='Option A' value='A' />
                        <option display='Option B' value='B' />
                        <option display='Option C' value='C' />
                    </options>
                    <selected>B</selected>
                </root>";

            // Add the XML as a custom XML part to the document.
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlContent);

            // Parse the XML to retrieve the option elements.
            XDocument xDoc = XDocument.Parse(xmlContent);
            var optionElements = xDoc.Root?
                .Element("options")?
                .Elements("option");

            // Create an inline drop‑down list content control.
            StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
            {
                Title = "DynamicDropdown",
                Tag = "dynamic-dropdown"
            };

            // Populate the drop‑down list items from the XML.
            if (optionElements != null)
            {
                foreach (var opt in optionElements)
                {
                    string display = (string)opt.Attribute("display") ?? string.Empty;
                    string value = (string)opt.Attribute("value") ?? string.Empty;
                    dropdown.ListItems.Add(new SdtListItem(display, value));
                }
            }

            // Bind the selected value of the drop‑down to the <selected> element in the XML part.
            dropdown.XmlMapping.SetMapping(xmlPart, "/root[1]/selected[1]", string.Empty);

            // Insert the content control into the first paragraph of the document.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.AppendChild(dropdown);

            // Save the resulting document.
            doc.Save("DropdownBound.docx");
        }
    }
}
