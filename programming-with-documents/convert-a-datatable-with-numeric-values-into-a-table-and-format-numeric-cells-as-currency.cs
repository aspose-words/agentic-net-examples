using System;
using System.Data;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with sample numeric data.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Price", typeof(decimal));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("Total", typeof(decimal));

        // Populate the DataTable.
        table.Rows.Add("Apple", 0.5m, 10, 5.0m);
        table.Rows.Add("Banana", 0.3m, 20, 6.0m);
        table.Rows.Add("Orange", 0.8m, 5, 4.0m);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table wordTable = builder.StartTable();

        // Insert header row.
        foreach (DataColumn col in table.Columns)
        {
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Bold = true;
            builder.Write(col.ColumnName);
        }
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in table.Rows)
        {
            foreach (DataColumn col in table.Columns)
            {
                builder.InsertCell();

                object value = row[col];
                // Format numeric values as currency.
                if (col.DataType == typeof(decimal) || col.DataType == typeof(double) || col.DataType == typeof(float))
                {
                    builder.Write(Convert.ToDecimal(value).ToString("C", CultureInfo.CurrentCulture));
                }
                else if (col.DataType == typeof(int) || col.DataType == typeof(long))
                {
                    // For integer quantities, display as plain number.
                    builder.Write(value.ToString());
                }
                else
                {
                    builder.Write(value.ToString());
                }
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
