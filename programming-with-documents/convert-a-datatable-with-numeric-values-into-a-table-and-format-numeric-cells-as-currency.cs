using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with sample data.
        DataTable data = new DataTable("Products");
        data.Columns.Add("Item", typeof(string));
        data.Columns.Add("Price", typeof(decimal));

        data.Rows.Add("Apple", 1.25m);
        data.Rows.Add("Banana", 0.75m);
        data.Rows.Add("Cherry", 2.50m);
        data.Rows.Add("Date", 3.10m);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Populate the table with DataTable rows.
        foreach (DataRow row in data.Rows)
        {
            // Item cell.
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            // Price cell – insert a field with currency formatting.
            builder.InsertCell();
            decimal price = (decimal)row["Price"];
            // The field calculates the value and formats it as currency.
            Field field = builder.InsertField($"= {price} \\# \"$#,##0.00\"");
            field.Update();

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableFromDataTable.docx");
    }
}
