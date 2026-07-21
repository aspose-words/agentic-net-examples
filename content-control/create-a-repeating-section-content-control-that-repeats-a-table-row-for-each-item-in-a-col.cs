using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "ItemsRepeatingSection",
            Tag = "items-repeating"
        };

        // Build a table with a header row and a placeholder row that will be repeated.
        Table table = new Table(doc);

        // Header row.
        Row headerRow = new Row(doc);
        Cell headerCell = new Cell(doc);
        headerCell.AppendChild(new Paragraph(doc));
        headerCell.FirstParagraph.AppendChild(new Run(doc, "Item"));
        headerRow.AppendChild(headerCell);
        table.AppendChild(headerRow);

        // Placeholder row (template for each item).
        Row placeholderRow = new Row(doc);
        Cell placeholderCell = new Cell(doc);
        placeholderCell.AppendChild(new Paragraph(doc));
        placeholderCell.FirstParagraph.AppendChild(new Run(doc, "{{Item}}"));
        placeholderRow.AppendChild(placeholderCell);
        table.AppendChild(placeholderRow);

        // Add the table to the repeating section.
        repeatingSection.AppendChild(table);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Sample collection of items to repeat.
        List<string> items = new List<string> { "Apple", "Banana", "Cherry" };

        // Locate the placeholder row inside the repeating section.
        Table repeatingTable = repeatingSection.GetChildNodes(NodeType.Table, true)
            .OfType<Table>()
            .FirstOrDefault() ?? throw new InvalidOperationException("Repeating table not found.");

        Row templateRow = repeatingTable.Rows[1]; // Index 0 = header, 1 = placeholder.
        int insertIndex = 1; // Position after header row.

        // Clone the placeholder row for each item and replace the placeholder text.
        foreach (string item in items)
        {
            Row newRow = (Row)templateRow.Clone(true);
            Cell cell = newRow.FirstCell;
            cell.RemoveAllChildren();

            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, item));
            cell.AppendChild(para);

            repeatingTable.Rows.Insert(insertIndex, newRow);
            insertIndex++;
        }

        // Remove the original placeholder row.
        templateRow.Remove();

        // Save the resulting document.
        string docPath = "RepeatingSectionTable.docx";
        doc.Save(docPath);

        // Serialize the items collection to JSON for demonstration.
        string json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText("items.json", json);
    }
}
