using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

public class Program
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
        Cell headerCell = new Cell(doc);
        header.AppendChild(headerCell);
        Paragraph headerPara = new Paragraph(doc);
        headerCell.AppendChild(headerPara);
        headerPara.AppendChild(new Run(doc, "Item"));

        // Create a repeating section content control at the row level.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc, SdtType.RepeatingSection, MarkupLevel.Row);
        table.AppendChild(repeatingSection);

        // Sample collection to repeat.
        string[] items = { "Apple", "Banana", "Cherry" };

        // For each item, create a repeating section item that contains a table row.
        foreach (string item in items)
        {
            // Repeating section item (row level).
            StructuredDocumentTag itemSdt = new StructuredDocumentTag(
                doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSection.AppendChild(itemSdt);

            // Row that will be repeated.
            Row row = new Row(doc);
            itemSdt.AppendChild(row);

            // Single cell with the item text.
            Cell cell = new Cell(doc);
            row.AppendChild(cell);
            Paragraph para = new Paragraph(doc);
            cell.AppendChild(para);
            para.AppendChild(new Run(doc, item));
        }

        // Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");
    }
}
