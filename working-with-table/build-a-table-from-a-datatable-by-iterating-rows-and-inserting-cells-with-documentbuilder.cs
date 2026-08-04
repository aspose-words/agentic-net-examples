using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableFromDataTable
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample DataTable with some data.
            DataTable dataTable = new DataTable("Sample");
            dataTable.Columns.Add("Product");
            dataTable.Columns.Add("Quantity");
            dataTable.Columns.Add("Price");
            dataTable.Rows.Add("Apples", 10, 1.5);
            dataTable.Rows.Add("Bananas", 20, 0.8);
            dataTable.Rows.Add("Carrots", 15, 0.6);

            // Initialize a new blank document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Add a header row using the column names.
            foreach (DataColumn column in dataTable.Columns)
            {
                builder.InsertCell();
                builder.Write(column.ColumnName);
            }
            builder.EndRow();

            // Populate the table with the rows from the DataTable.
            foreach (DataRow dataRow in dataTable.Rows)
            {
                foreach (object value in dataRow.ItemArray)
                {
                    builder.InsertCell();
                    builder.Write(value?.ToString() ?? string.Empty);
                }
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("TableFromDataTable.docx");
        }
    }
}
