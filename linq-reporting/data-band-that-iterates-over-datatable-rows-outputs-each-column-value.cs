using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Prepare a sample DataTable with some columns and rows.
        DataTable dataTable = new DataTable("Sample");
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));

        dataTable.Rows.Add(1, "Apple", 0.5m);
        dataTable.Rows.Add(2, "Banana", 0.3m);
        dataTable.Rows.Add(3, "Cherry", 1.2m);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table in the document.
        Table wordTable = builder.StartTable();

        // Insert a header row containing the column names.
        foreach (DataColumn column in dataTable.Columns)
        {
            builder.InsertCell();
            builder.Write(column.ColumnName);
        }
        builder.EndRow();

        // Iterate over each DataRow and output its column values into the table.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object value in row.ItemArray)
            {
                builder.InsertCell();
                builder.Write(value?.ToString() ?? string.Empty);
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("DataBand.docx");
    }
}
