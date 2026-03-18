using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Build a table with a header row (Product | Price).
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // 3. Add a custom XML part that holds the collection we want to repeat.
        //    The XML structure contains a <products> root with multiple <product> elements.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(
            "Products",
            "<products>" +
                "<product>" +
                    "<name>Apple</name>" +
                    "<price>1.20</price>" +
                "</product>" +
                "<product>" +
                    "<name>Banana</name>" +
                    "<price>0.80</price>" +
                "</product>" +
                "<product>" +
                    "<name>Cherry</name>" +
                    "<price>2.50</price>" +
                "</product>" +
            "</products>");

        // 4. Insert a RepeatingSection content control at the row level.
        //    This control will repeat for each <product> node selected by the XPath.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,
            MarkupLevel.Row);

        // Map the repeating section to the collection of <product> nodes.
        repeatingSection.XmlMapping.SetMapping(xmlPart, "/products[1]/product", string.Empty);

        // Append the repeating section to the table (it becomes a child of the table).
        table.AppendChild(repeatingSection);

        // 5. Inside the repeating section, add a RepeatingSectionItem.
        //    Each item represents one iteration (i.e., one table row).
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSectionItem,
            MarkupLevel.Row);
        repeatingSection.AppendChild(repeatingItem);

        // 6. Create the row that will be duplicated for each item.
        Row dataRow = new Row(doc);
        repeatingItem.AppendChild(dataRow);

        // 7. First cell – bind to the <name> element of the current <product>.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,
            MarkupLevel.Cell);
        nameSdt.XmlMapping.SetMapping(xmlPart, "/products[1]/product[1]/name[1]", string.Empty);
        dataRow.AppendChild(nameSdt);

        // 8. Second cell – bind to the <price> element of the current <product>.
        StructuredDocumentTag priceSdt = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,
            MarkupLevel.Cell);
        priceSdt.XmlMapping.SetMapping(xmlPart, "/products[1]/product[1]/price[1]", string.Empty);
        dataRow.AppendChild(priceSdt);

        // 9. Close the table construction.
        builder.EndTable();

        // 10. Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");
    }
}
