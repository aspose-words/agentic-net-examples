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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Insert the table into the document.
        Table table = builder.EndTable();

        // Create a repeating section content control at the row level.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
        table.AppendChild(repeatingSection);

        // Sample data collection.
        List<Item> items = new List<Item>
        {
            new Item { Name = "Apple", Quantity = 10 },
            new Item { Name = "Banana", Quantity = 5 },
            new Item { Name = "Cherry", Quantity = 20 }
        };

        // For each item, create a repeating section item that contains a table row.
        foreach (Item item in items)
        {
            // Create a repeating section item (row level).
            StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSection.AppendChild(repeatingItem);

            // Create a new row for this item.
            Row row = new Row(doc);
            repeatingItem.AppendChild(row);

            // First cell – product name.
            Cell nameCell = new Cell(doc);
            row.AppendChild(nameCell);
            Paragraph namePara = new Paragraph(doc);
            nameCell.AppendChild(namePara);
            namePara.AppendChild(new Run(doc, item.Name));

            // Second cell – quantity.
            Cell qtyCell = new Cell(doc);
            row.AppendChild(qtyCell);
            Paragraph qtyPara = new Paragraph(doc);
            qtyCell.AppendChild(qtyPara);
            qtyPara.AppendChild(new Run(doc, item.Quantity.ToString()));
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RepeatingSectionTable.docx");
        doc.Save(outputPath);
    }

    // Simple data model for demonstration.
    private class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }
}
