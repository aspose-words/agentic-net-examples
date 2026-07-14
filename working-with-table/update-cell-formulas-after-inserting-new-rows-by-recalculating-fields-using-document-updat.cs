using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a formula field that sums the values above it.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Formula row – the field will calculate the sum of the values above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a formula field. The field code does not include the surrounding braces.
        builder.InsertField("=SUM(ABOVE)");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Insert a new data row before the formula row.
        // The formula row is currently the last row in the table.
        int formulaRowIndex = table.Rows.Count - 1;

        // Create a new row and insert it at the calculated position.
        Row newRow = new Row(doc);
        table.Rows.Insert(formulaRowIndex, newRow);

        // Populate the new row with cells.
        Cell itemCell = new Cell(doc);
        itemCell.AppendChild(new Paragraph(doc));
        itemCell.FirstParagraph.AppendChild(new Run(doc, "Orange"));
        newRow.AppendChild(itemCell);

        Cell valueCell = new Cell(doc);
        valueCell.AppendChild(new Paragraph(doc));
        valueCell.FirstParagraph.AppendChild(new Run(doc, "30"));
        newRow.AppendChild(valueCell);

        // Recalculate all fields in the document, including the formula field.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("UpdatedTableFields.docx");
    }
}
