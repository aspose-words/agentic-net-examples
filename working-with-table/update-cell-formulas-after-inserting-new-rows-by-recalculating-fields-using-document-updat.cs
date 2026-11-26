using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "UpdatedTable.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header, two data rows, and a summary row that contains a formula.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Row1");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row2");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Summary row with a formula that sums the values above.
        builder.InsertCell();
        builder.Write("Sum");
        builder.InsertCell();
        // Insert a field that will calculate the sum of the "Value" column.
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Insert a new data row after the first data row (index 2 in the Rows collection).
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell newCell1 = new Cell(doc);
        newCell1.AppendChild(new Paragraph(doc));
        newCell1.FirstParagraph.AppendChild(new Run(doc, "Row1.5"));
        newRow.Cells.Add(newCell1);

        // Second cell of the new row with a numeric value.
        Cell newCell2 = new Cell(doc);
        newCell2.AppendChild(new Paragraph(doc));
        newCell2.FirstParagraph.AppendChild(new Run(doc, "30"));
        newRow.Cells.Add(newCell2);

        // Insert the new row into the table.
        table.Rows.Insert(2, newRow);

        // Recalculate all fields in the document, updating the sum formula.
        doc.UpdateFields();

        // Retrieve the updated sum from the summary row.
        Row summaryRow = table.Rows[table.Rows.Count - 1];
        string sumResult = summaryRow.Cells[1].GetText().Trim();

        // Output the result to the console.
        Console.WriteLine($"Sum after inserting new row and updating fields: {sumResult}");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
