using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "ItemsRepeatingSection";
        repeatingSection.Tag = "items-section";

        // Create a table that will hold the rows.
        Table table = new Table(doc);

        // Create a template row with placeholder text.
        Row templateRow = new Row(doc);

        // First cell – placeholder for the item name.
        Cell cell1 = new Cell(doc);
        Paragraph para1 = new Paragraph(doc);
        Run run1 = new Run(doc, "{Item}");
        para1.AppendChild(run1);
        cell1.AppendChild(para1);
        templateRow.AppendChild(cell1);

        // Second cell – static description text.
        Cell cell2 = new Cell(doc);
        Paragraph para2 = new Paragraph(doc);
        Run run2 = new Run(doc, "Description");
        para2.AppendChild(run2);
        cell2.AppendChild(para2);
        templateRow.AppendChild(cell2);

        // Add the template row to the table.
        table.AppendChild(templateRow);

        // Add the table to the repeating section.
        repeatingSection.AppendChild(table);

        // Add the repeating section to the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Sample collection of items.
        List<string> items = new List<string> { "Apple", "Banana", "Cherry" };

        // For each item, clone the template row, replace the placeholder, and add to the table.
        foreach (string item in items)
        {
            Row newRow = (Row)templateRow.Clone(true);
            // Replace placeholder text in the first cell.
            Cell firstCell = newRow.Cells[0];
            if (firstCell.FirstParagraph != null && firstCell.FirstParagraph.Runs.Count > 0)
            {
                Run placeholderRun = firstCell.FirstParagraph.Runs[0];
                placeholderRun.Text = item;
            }

            table.AppendChild(newRow);
        }

        // Remove the original template row that contains the placeholder.
        table.Rows.Remove(templateRow);

        // Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");

        // Serialize the collection to JSON as an additional output.
        string json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText("items.json", json);
    }
}
