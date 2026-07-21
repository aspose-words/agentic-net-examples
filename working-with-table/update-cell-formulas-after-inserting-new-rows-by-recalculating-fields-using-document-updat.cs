using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header, two data rows, and a total row that contains a formula field.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Total row with a formula field that sums the values above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert the formula field. The field code "=SUM(ABOVE)" will sum the numeric cells in the column.
        builder.InsertField(" =SUM(ABOVE) ");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Insert a new data row before the total row.
        // The total row is currently the last row (index 3). Insert at index 3 to place the new row before it.
        Row newRow = new Row(doc);

        // Cell for the item name.
        Cell itemCell = new Cell(doc);
        itemCell.AppendChild(new Paragraph(doc));
        itemCell.FirstParagraph.AppendChild(new Run(doc, "Oranges"));
        newRow.AppendChild(itemCell);

        // Cell for the quantity.
        Cell qtyCell = new Cell(doc);
        qtyCell.AppendChild(new Paragraph(doc));
        qtyCell.FirstParagraph.AppendChild(new Run(doc, "15"));
        newRow.AppendChild(qtyCell);

        // Insert the new row into the table.
        table.Rows.Insert(3, newRow);

        // Recalculate all fields in the document, including the formula field.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableWithFormula.docx");
        doc.Save(outputPath);
    }
}
