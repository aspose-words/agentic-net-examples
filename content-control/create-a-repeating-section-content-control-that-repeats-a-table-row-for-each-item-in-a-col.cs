using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

public class RepeatingSectionExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build a simple table with a header row.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Header row.
        Row header = new Row(doc);
        table.AppendChild(header);

        Cell headerCell1 = new Cell(doc);
        header.AppendChild(headerCell1);
        Paragraph headerPara1 = new Paragraph(doc);
        headerCell1.AppendChild(headerPara1);
        headerPara1.AppendChild(new Run(doc, "Name"));

        Cell headerCell2 = new Cell(doc);
        header.AppendChild(headerCell2);
        Paragraph headerPara2 = new Paragraph(doc);
        headerCell2.AppendChild(headerPara2);
        headerPara2.AppendChild(new Run(doc, "Value"));

        // Create a repeating section content control at the row level.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        repeatingSection.Title = "Items";
        table.AppendChild(repeatingSection);

        // Sample data collection.
        var items = new List<(string Name, string Value)>
        {
            ("Alice", "10"),
            ("Bob", "20"),
            ("Charlie", "30")
        };

        // For each item, add a repeating section item that contains a table row.
        foreach (var item in items)
        {
            // Repeating section item (row level).
            StructuredDocumentTag itemSdt = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSection.AppendChild(itemSdt);

            // The actual row that will be repeated.
            Row dataRow = new Row(doc);
            itemSdt.AppendChild(dataRow);

            // First cell – Name.
            Cell nameCell = new Cell(doc);
            dataRow.AppendChild(nameCell);
            Paragraph namePara = new Paragraph(doc);
            nameCell.AppendChild(namePara);
            namePara.AppendChild(new Run(doc, item.Name));

            // Second cell – Value.
            Cell valueCell = new Cell(doc);
            dataRow.AppendChild(valueCell);
            Paragraph valuePara = new Paragraph(doc);
            valueCell.AppendChild(valuePara);
            valuePara.AppendChild(new Run(doc, item.Value));
        }

        // Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");
    }
}
