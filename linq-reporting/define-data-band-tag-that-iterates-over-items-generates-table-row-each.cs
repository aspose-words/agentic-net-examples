using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class DataBandExample
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder to work with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Add a CustomXmlPart that will serve as the data source for the band.
        // -----------------------------------------------------------------
        // The XML contains a collection of <item> elements, each with <name> and <price>.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add("Items",
            "<items>" +
                "<item>" +
                    "<name>Apple</name>" +
                    "<price>1.20</price>" +
                "</item>" +
                "<item>" +
                    "<name>Banana</name>" +
                    "<price>0.80</price>" +
                "</item>" +
                "<item>" +
                    "<name>Cherry</name>" +
                    "<price>2.50</price>" +
                "</item>" +
            "</items>");

        // -----------------------------------------------------------------
        // 2. Build a table with a header row (Name | Price).
        // -----------------------------------------------------------------
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();   // End of header row.
        builder.EndTable(); // Finish the table so we can append the repeating section.

        // -----------------------------------------------------------------
        // 3. Create a RepeatingSection StructuredDocumentTag at the row level.
        //    This tag will repeat for each <item> node in the XML.
        // -----------------------------------------------------------------
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        // Map the repeating section to the collection of <item> elements.
        repeatingSection.XmlMapping.SetMapping(xmlPart, "/items[1]/item", string.Empty);
        table.AppendChild(repeatingSection);

        // -----------------------------------------------------------------
        // 4. Inside the repeating section, create a RepeatingSectionItem tag.
        //    This represents a single row that will be cloned for each item.
        // -----------------------------------------------------------------
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
        repeatingSection.AppendChild(repeatingItem);

        // -----------------------------------------------------------------
        // 5. Create the row that will be repeated.
        // -----------------------------------------------------------------
        Row dataRow = new Row(doc);
        repeatingItem.AppendChild(dataRow);

        // -----------------------------------------------------------------
        // 6. Add cells with PlainText StructuredDocumentTags bound to the XML fields.
        // -----------------------------------------------------------------
        // Cell for the <name> element.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        nameSdt.XmlMapping.SetMapping(xmlPart, "/items[1]/item[1]/name[1]", string.Empty);
        dataRow.AppendChild(nameSdt);

        // Cell for the <price> element.
        StructuredDocumentTag priceSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        priceSdt.XmlMapping.SetMapping(xmlPart, "/items[1]/item[1]/price[1]", string.Empty);
        dataRow.AppendChild(priceSdt);

        // -----------------------------------------------------------------
        // 7. Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save("DataBandTable.docx");
    }
}
