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

        // Build a table with two numeric cells and a formula field that sums the numbers above.
        builder.StartTable();

        // First row, first cell: numeric value 5
        builder.InsertCell();
        builder.Write("5");
        builder.EndRow();

        // Second row, first cell: numeric value 15
        builder.InsertCell();
        builder.Write("15");
        builder.EndRow();

        // Third row, first cell: formula field =SUM(ABOVE)
        builder.InsertCell();
        // Insert a formula field that sums the numbers in the column above.
        builder.InsertField("=SUM(ABOVE)", "?");
        builder.EndRow();

        builder.EndTable();

        // Save the initial document.
        string initialPath = "TableWithFormula.docx";
        doc.Save(initialPath);

        // Modify the first numeric cell (change 5 to 20).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = table.Rows[0].Cells[0];
        // Clear existing content and write new value.
        firstCell.RemoveAllChildren();
        firstCell.AppendChild(new Paragraph(doc));
        firstCell.FirstParagraph.AppendChild(new Run(doc, "20"));

        // Recalculate fields after the modification.
        doc.UpdateFields();

        // Save the updated document.
        string updatedPath = "TableWithFormulaUpdated.docx";
        doc.Save(updatedPath);

        // Verify that the output files were created.
        if (!File.Exists(initialPath) || !File.Exists(updatedPath))
        {
            throw new InvalidOperationException("One or more output files were not created successfully.");
        }
    }
}
