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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row: headers.
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second row: total with a formula field that sums the values above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a formula field. This overload inserts the field and updates it automatically.
        builder.InsertField("= SUM(ABOVE) \\# #,##0");
        builder.EndRow();

        builder.EndTable();

        // Save the initial document.
        string initialPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithField.docx");
        doc.Save(initialPath);

        // Modify the quantity value in the first row, second column.
        Cell quantityCell = table.Rows[0].Cells[1];
        // Clear existing runs and insert new text.
        quantityCell.FirstParagraph.Runs.Clear();
        quantityCell.FirstParagraph.AppendChild(new Run(doc, "25"));

        // Recalculate all fields in the document.
        doc.UpdateFields();

        // Save the updated document.
        string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithField_Updated.docx");
        doc.Save(updatedPath);

        // Verify that the updated file was created.
        if (!File.Exists(updatedPath))
            throw new InvalidOperationException("The updated document was not saved correctly.");
    }
}
