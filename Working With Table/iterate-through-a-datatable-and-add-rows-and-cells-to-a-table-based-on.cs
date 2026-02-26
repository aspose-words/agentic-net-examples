using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

class TableFromDataTable
{
    static void Main()
    {
        // Create a sample DataTable with some data.
        DataTable data = new DataTable("Sample");
        data.Columns.Add("Product");
        data.Columns.Add("Quantity");
        data.Columns.Add("Price");

        data.Rows.Add("Apples", 10, 0.5);
        data.Rows.Add("Bananas", 5, 0.3);
        data.Rows.Add("Cherries", 20, 1.2);

        // Initialize a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---- Header Row -------------------------------------------------
        // Insert a cell for each column name and write the header text.
        foreach (DataColumn col in data.Columns)
        {
            builder.InsertCell();
            builder.Write(col.ColumnName);
        }
        // End the header row.
        builder.EndRow();

        // ---- Data Rows --------------------------------------------------
        // Iterate through each DataRow in the DataTable.
        foreach (DataRow row in data.Rows)
        {
            // For each column in the current row, insert a cell and write its value.
            foreach (DataColumn col in data.Columns)
            {
                builder.InsertCell();
                // Convert the cell value to string, handling nulls safely.
                string cellText = row[col] != DBNull.Value ? row[col].ToString() : string.Empty;
                builder.Write(cellText);
            }
            // End the current row before moving to the next one.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Optional: Auto‑fit the table to its contents for a cleaner layout.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        doc.Save("TableFromDataTable.docx");
    }
}
