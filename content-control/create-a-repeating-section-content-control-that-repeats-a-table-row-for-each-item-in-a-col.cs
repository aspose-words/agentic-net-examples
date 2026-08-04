using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace RepeatingSectionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with a header row.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Retrieve the Table that the builder is currently working on.
            // The current paragraph resides inside a cell; get the ancestor Table.
            Table table = (Table)builder.CurrentParagraph.GetAncestor(NodeType.Table);

            // Insert a repeating section content control at the row level.
            StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
            table.AppendChild(repeatingSection);

            // Inside the repeating section, insert a repeating section item (also at row level).
            StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSection.AppendChild(repeatingItem);

            // Create a template row that will be cloned for each data item.
            Row templateRow = new Row(doc);
            repeatingItem.AppendChild(templateRow);

            // First cell – will hold the item name.
            Cell cellItem = new Cell(doc);
            cellItem.AppendChild(new Paragraph(doc));
            cellItem.FirstParagraph.AppendChild(new Run(doc, string.Empty));
            templateRow.AppendChild(cellItem);

            // Second cell – will hold the quantity.
            Cell cellQty = new Cell(doc);
            cellQty.AppendChild(new Paragraph(doc));
            cellQty.FirstParagraph.AppendChild(new Run(doc, string.Empty));
            templateRow.AppendChild(cellQty);

            // Sample data collection.
            List<(string Item, string Quantity)> data = new List<(string, string)>
            {
                ("Apple", "10"),
                ("Banana", "20"),
                ("Cherry", "30")
            };

            // Clone the template row for each data entry and fill the cells.
            foreach (var entry in data)
            {
                // Clone the repeating item (which contains the template row).
                StructuredDocumentTag clonedItem = (StructuredDocumentTag)repeatingItem.Clone(true);

                // Locate the row inside the cloned item.
                Row clonedRow = clonedItem.GetChildNodes(NodeType.Row, true).Cast<Row>().First();

                // Fill the first cell.
                Cell clonedCellItem = clonedRow.FirstCell;
                clonedCellItem.FirstParagraph.Runs[0].Text = entry.Item;

                // Fill the second cell.
                Cell clonedCellQty = clonedRow.LastCell;
                clonedCellQty.FirstParagraph.Runs[0].Text = entry.Quantity;

                // Append the populated item to the repeating section.
                repeatingSection.AppendChild(clonedItem);
            }

            // Remove the original template item – it was only a placeholder.
            repeatingItem.Remove();

            // Save the resulting document.
            string docPath = Path.Combine(Environment.CurrentDirectory, "RepeatingSectionTable.docx");
            doc.Save(docPath);

            // Serialize information about the repeating sections to JSON.
            var payload = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
                .Select(sdt => new
                {
                    Title = sdt.Title,
                    Tag = sdt.Tag,
                    Text = sdt.GetText().Trim()
                })
                .ToList();

            string json = JsonConvert.SerializeObject(payload, Formatting.Indented);
            string jsonPath = Path.Combine(Environment.CurrentDirectory, "RepeatingSectionInfo.json");
            File.WriteAllText(jsonPath, json);
        }
    }
}
