using System;
using System.IO;
using System.Xml;
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

            // Define XML data that will be stored in a custom XML part.
            // It contains a <selected> element that will hold the chosen value
            // and a list of <option> elements that will be used to fill the dropdown.
            string xmlContent = @"
                <root>
                    <selected></selected>
                    <options>
                        <option value='A'>Option A</option>
                        <option value='B'>Option B</option>
                        <option value='C'>Option C</option>
                    </options>
                </root>";

            // Add the XML as a custom XML part to the document.
            string partId = Guid.NewGuid().ToString("B");
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xmlContent);

            // Create a dropdown list content control (inline level).
            StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
            {
                Title = "SampleDropdown",
                Tag = "sample-dropdown"
            };

            // Populate the dropdown list items from the XML part.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlContent);
            XmlNodeList optionNodes = xmlDoc.SelectNodes("/root/options/option");
            if (optionNodes != null)
            {
                foreach (XmlNode node in optionNodes)
                {
                    string displayText = node.InnerText ?? string.Empty;
                    string value = node.Attributes?["value"]?.Value ?? displayText;
                    dropdown.ListItems.Add(new SdtListItem(displayText, value));
                }
            }

            // Bind the selected value of the dropdown to the <selected> element in the XML part.
            dropdown.XmlMapping.SetMapping(xmlPart, "/root[1]/selected[1]", string.Empty);

            // Insert the dropdown into the first paragraph of the document.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            paragraph.AppendChild(dropdown);

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DropdownMapped.docx");
            doc.Save(outputPath);
        }
    }
}
