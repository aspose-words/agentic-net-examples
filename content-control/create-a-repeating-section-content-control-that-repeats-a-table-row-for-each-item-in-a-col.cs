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
        // Sample data collection.
        var items = new List<Item>
        {
            new Item { Name = "Apple", Quantity = 5 },
            new Item { Name = "Banana", Quantity = 12 },
            new Item { Name = "Cherry", Quantity = 20 }
        };

        // Create a new blank document.
        var doc = new Document();

        // Create a block‑level repeating section content control.
        var repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "ItemsRepeatingSection",
            Tag = "items-section"
        };
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Inside the repeating section, create a table that will hold the rows.
        var table = new Table(doc);
        repeatingSection.AppendChild(table);

        // Create a template row with placeholder text.
        var templateRow = new Row(doc);
        var nameCell = new Cell(doc);
        nameCell.AppendChild(new Paragraph(doc));
        nameCell.FirstParagraph.AppendChild(new Run(doc, "{Name}"));
        var qtyCell = new Cell(doc);
        qtyCell.AppendChild(new Paragraph(doc));
        qtyCell.FirstParagraph.AppendChild(new Run(doc, "{Quantity}"));
        templateRow.Cells.Add(nameCell);
        templateRow.Cells.Add(qtyCell);
        table.Rows.Add(templateRow);

        // For each item, clone the template row, replace placeholders, and add to the table.
        foreach (var item in items)
        {
            var newRow = (Row)templateRow.Clone(true);
            foreach (Cell cell in newRow.Cells)
            {
                foreach (Run run in cell.GetChildNodes(NodeType.Run, true).OfType<Run>())
                {
                    if (run.Text.Contains("{Name}"))
                        run.Text = run.Text.Replace("{Name}", item.Name);
                    if (run.Text.Contains("{Quantity}"))
                        run.Text = run.Text.Replace("{Quantity}", item.Quantity.ToString());
                }
            }
            table.Rows.Add(newRow);
        }

        // Remove the template row so only real data rows remain.
        table.Rows.Remove(templateRow);

        // Save the resulting document.
        doc.Save("RepeatingSectionTable.docx");

        // Optionally, serialize the collection to JSON for verification.
        var json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText("items.json", json);
    }

    // Simple data model.
    private class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }
}
