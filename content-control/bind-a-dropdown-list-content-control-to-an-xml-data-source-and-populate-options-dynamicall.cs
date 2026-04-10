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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a prompt for the user.
        builder.Writeln("Select an option:");

        // -----------------------------------------------------------------
        // 1. Create a custom XML part that holds the list of options.
        // -----------------------------------------------------------------
        string xmlContent =
            "<root>" +
                "<options>" +
                    "<option>Apple</option>" +
                    "<option>Banana</option>" +
                    "<option>Cherry</option>" +
                "</options>" +
            "</root>";

        string xmlPartId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // -----------------------------------------------------------------
        // 2. Insert a DropDownList content control (inline level).
        // -----------------------------------------------------------------
        StructuredDocumentTag dropDown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline);
        builder.InsertNode(dropDown);

        // -----------------------------------------------------------------
        // 3. Populate the dropdown list items dynamically from the XML part.
        // -----------------------------------------------------------------
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlContent);
        XmlNodeList optionNodes = xmlDoc.SelectNodes("/root/options/option");

        foreach (XmlNode node in optionNodes)
        {
            // Each option becomes a list item. Use the same text for display and value.
            dropDown.ListItems.Add(new SdtListItem(node.InnerText));
        }

        // -----------------------------------------------------------------
        // 4. (Optional) Bind the content control to the first option node.
        //    This demonstrates XML mapping for the selected value.
        // -----------------------------------------------------------------
        dropDown.XmlMapping.SetMapping(xmlPart, "/root/options[1]/option[1]", string.Empty);

        // Save the resulting document.
        doc.Save("DropDownListBinding.docx");
    }
}
