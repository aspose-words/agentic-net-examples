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

        // Build a simple table with a header, two data rows and a formula row that sums the values above.
        Table table = builder.StartTable();

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

        // Formula row – the second cell contains a SUM(ABOVE) field.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a formula field that will sum the numbers in the column above.
        // Use the string overload of InsertField which inserts the field and updates it immediately.
        builder.InsertField("=SUM(ABOVE)");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Insert a new data row between the existing data rows.
        // Create a new row with two cells.
        Row newRow = new Row(doc);
        Cell cellItem = new Cell(doc);
        cellItem.AppendChild(new Paragraph(doc));
        cellItem.FirstParagraph.AppendChild(new Run(doc, "Oranges"));
        newRow.AppendChild(cellItem);

        Cell cellQty = new Cell(doc);
        cellQty.AppendChild(new Paragraph(doc));
        cellQty.FirstParagraph.AppendChild(new Run(doc, "30"));
        newRow.AppendChild(cellQty);

        // Insert the new row before the formula row (which is currently at index 3).
        table.Rows.Insert(3, newRow);

        // Recalculate all fields in the document, including the formula field.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableFormulas.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
