using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class InTableListFromXml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Add a custom XML part that contains groups and items.
        // -----------------------------------------------------------------
        string xml = @"
<root>
    <group>
        <name>Group A</name>
        <item>Item A1</item>
        <item>Item A2</item>
        <item>Item A3</item>
    </group>
    <group>
        <name>Group B</name>
        <item>Item B1</item>
        <item>Item B2</item>
    </group>
    <group>
        <name>Group C</name>
        <item>Item C1</item>
    </group>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add("Groups", xml);

        // -----------------------------------------------------------------
        // 2. Build the outer table that will hold each group in a separate row.
        // -----------------------------------------------------------------
        Table outerTable = builder.StartTable();
        builder.InsertCell(); builder.Write("Group");
        builder.InsertCell(); builder.Write("Items");
        builder.EndRow();

        // -----------------------------------------------------------------
        // 3. Create a repeating section (row level) that repeats for each <group>.
        // -----------------------------------------------------------------
        StructuredDocumentTag groupRepeating = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        groupRepeating.XmlMapping.SetMapping(xmlPart, "/root[1]/group", string.Empty);
        outerTable.AppendChild(groupRepeating);

        // The repeating section needs a RepeatingSectionItem that represents one row.
        StructuredDocumentTag groupItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
        groupRepeating.AppendChild(groupItem);

        // Row that will be cloned for each group.
        Row groupRow = new Row(doc);
        groupItem.AppendChild(groupRow);

        // -------------------------------------------------------------
        // 3a. First cell – map the group name.
        // -------------------------------------------------------------
        Cell nameCell = new Cell(doc);
        groupRow.AppendChild(nameCell);
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        nameSdt.XmlMapping.SetMapping(xmlPart, "./name", string.Empty);
        nameCell.AppendChild(nameSdt);

        // -------------------------------------------------------------
        // 3b. Second cell – contains a nested table with the group's items.
        // -------------------------------------------------------------
        Cell itemsCell = new Cell(doc);
        groupRow.AppendChild(itemsCell);

        // Move the builder cursor into the new cell so we can start a nested table.
        builder.MoveTo(itemsCell.FirstParagraph);
        Table innerTable = builder.StartTable();
        builder.InsertCell(); builder.Write("Item");
        builder.EndRow();

        // -----------------------------------------------------------------
        // 4. Create a repeating section inside the inner table for each <item>.
        // -----------------------------------------------------------------
        StructuredDocumentTag itemRepeating = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        // Relative XPath – selects each <item> node that is a child of the current <group>.
        itemRepeating.XmlMapping.SetMapping(xmlPart, "./item", string.Empty);
        innerTable.AppendChild(itemRepeating);

        // RepeatingSectionItem that will be cloned for each item.
        StructuredDocumentTag itemItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
        itemRepeating.AppendChild(itemItem);

        // Row that will hold the item text.
        Row itemRow = new Row(doc);
        itemItem.AppendChild(itemRow);
        Cell itemCell = new Cell(doc);
        itemRow.AppendChild(itemCell);

        // PlainText SDT that maps to the current <item> node's text.
        StructuredDocumentTag itemSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
        // "." means the current node (the <item> element) – its inner text will be used.
        itemSdt.XmlMapping.SetMapping(xmlPart, ".", string.Empty);
        itemCell.AppendChild(itemSdt);

        // -----------------------------------------------------------------
        // 5. Apply a list template to the item rows so they appear as a bulleted list.
        // -----------------------------------------------------------------
        // Create a bullet list based on a predefined template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Move the builder to the paragraph inside the item cell and apply the list.
        builder.MoveTo(itemCell.FirstParagraph);
        builder.ListFormat.List = bulletList;
        // The paragraph itself will be empty now; when the XML is merged the text will appear as a list item.
        builder.Writeln(string.Empty);
        // Reset the builder's list format to avoid affecting later content.
        builder.ListFormat.RemoveNumbers();

        // End the inner and outer tables.
        builder.EndTable(); // ends innerTable
        builder.EndTable(); // ends outerTable

        // -----------------------------------------------------------------
        // 6. Save the document.
        // -----------------------------------------------------------------
        doc.Save("InTableListFromGroupedXml.docx");
    }
}
