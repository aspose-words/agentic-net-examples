using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for building content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two columns: Item and Quantity.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Writeln("A");
        builder.InsertCell();
        builder.Writeln("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln("B");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Formula row that sums the values above in the Quantity column.
        builder.InsertCell();
        builder.Writeln("Total");
        builder.InsertCell();
        // Insert a formula field =SUM(ABOVE). The field result is a placeholder ("0").
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the initial document (optional, demonstrates the state before insertion).
        string initialPath = "TableWithFormula.docx";
        doc.Save(initialPath);

        // Locate the table we just created.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not found in the document.");

        // Clone the structure of an existing row (the first data row) to create a new row.
        Row templateRow = table.Rows[1]; // Row with data (index 1 after header).
        Row newRow = (Row)templateRow.Clone(true);

        // Set the cell values for the new row.
        // First cell: Item name "C".
        Cell itemCell = newRow.Cells[0];
        itemCell.RemoveAllChildren();
        Paragraph pItem = new Paragraph(doc);
        pItem.AppendChild(new Run(doc, "C"));
        itemCell.AppendChild(pItem);

        // Second cell: Quantity "30".
        Cell qtyCell = newRow.Cells[1];
        qtyCell.RemoveAllChildren();
        Paragraph pQty = new Paragraph(doc);
        pQty.AppendChild(new Run(doc, "30"));
        qtyCell.AppendChild(pQty);

        // Insert the new row just before the formula row (which is the last row).
        int formulaRowIndex = table.Rows.Count - 1;
        table.Rows.Insert(formulaRowIndex, newRow);

        // Recalculate all fields in the document, including the formula field.
        doc.UpdateFields();

        // Save the updated document.
        string outputPath = "TableWithUpdatedFormula.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
