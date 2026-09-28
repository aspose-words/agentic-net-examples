using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x3 table with a formula field that sums the numbers above.
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
        builder.Write("Oranges");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Sum row with a formula field =SUM(ABOVE) in the second column.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null);
        builder.EndRow();

        builder.EndTable();

        // Save the initial document.
        string initialPath = "TableWithFormula.docx";
        doc.Save(initialPath);

        // Modify a cell value (change quantity of Apples from 10 to 15).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        Table table = (Table)tables[0];
        Row firstDataRow = table.Rows[1]; // Row after header.
        Cell quantityCell = firstDataRow.Cells[1];
        // Clear existing runs and insert new text.
        quantityCell.FirstParagraph.Runs.Clear();
        quantityCell.FirstParagraph.AppendChild(new Run(doc, "15"));

        // Recalculate all fields in the document.
        doc.UpdateFields();

        // Save the updated document.
        string updatedPath = "UpdatedTableFields.docx";
        doc.Save(updatedPath);

        // Verify that the updated file was saved.
        if (!File.Exists(updatedPath))
        {
            throw new Exception("The updated document was not saved correctly.");
        }
    }
}
