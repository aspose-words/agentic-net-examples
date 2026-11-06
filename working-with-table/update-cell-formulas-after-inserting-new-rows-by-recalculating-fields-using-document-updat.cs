using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header, two data rows, and a total row that uses a formula.
        builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Write("Item");
        builder.InsertCell(); builder.Write("Quantity");
        builder.InsertCell(); builder.Write("Price");
        builder.EndRow();

        // First data row.
        builder.InsertCell(); builder.Write("Apple");
        builder.InsertCell(); builder.Write("2");
        builder.InsertCell(); builder.Write("3");
        builder.EndRow();

        // Second data row.
        builder.InsertCell(); builder.Write("Banana");
        builder.InsertCell(); builder.Write("5");
        builder.InsertCell(); builder.Write("1");
        builder.EndRow();

        // Total row – the Quantity cell contains a formula that sums the values above.
        builder.InsertCell(); builder.Write("Total");
        builder.InsertCell(); builder.InsertField("=SUM(ABOVE)", null);
        builder.InsertCell(); builder.Write(string.Empty);
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new Exception("Table was not created.");

        // Insert a new data row before the total row.
        Row newRow = new Row(doc);

        // Item cell.
        Cell itemCell = new Cell(doc);
        itemCell.AppendChild(new Paragraph(doc));               // Ensure a paragraph exists.
        itemCell.FirstParagraph.AppendChild(new Run(doc, "Orange"));
        newRow.AppendChild(itemCell);

        // Quantity cell.
        Cell quantityCell = new Cell(doc);
        quantityCell.AppendChild(new Paragraph(doc));
        quantityCell.FirstParagraph.AppendChild(new Run(doc, "3"));
        newRow.AppendChild(quantityCell);

        // Price cell.
        Cell priceCell = new Cell(doc);
        priceCell.AppendChild(new Paragraph(doc));
        priceCell.FirstParagraph.AppendChild(new Run(doc, "2"));
        newRow.AppendChild(priceCell);

        // Insert the new row just before the total row (which is the last row).
        table.Rows.Insert(table.Rows.Count - 1, newRow);

        // Recalculate all fields (the SUM formula) after inserting the new row.
        doc.UpdateFields();

        // Save the document.
        const string outputPath = "UpdatedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
