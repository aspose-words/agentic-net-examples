using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFromDataTable
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare a DataTable with sample data.
            // ------------------------------------------------------------
            DataTable data = new DataTable("Sample");
            data.Columns.Add("Product", typeof(string));
            data.Columns.Add("Quantity", typeof(int));
            data.Columns.Add("Price", typeof(decimal));

            data.Rows.Add("Apples", 10, 0.5m);
            data.Rows.Add("Bananas", 20, 0.3m);
            data.Rows.Add("Carrots", 15, 0.2m);

            // ------------------------------------------------------------
            // 2. Create a new blank document and associate a DocumentBuilder.
            // ------------------------------------------------------------
            Document doc = new Document();                     // create document
            DocumentBuilder builder = new DocumentBuilder(doc); // create builder

            // ------------------------------------------------------------
            // 3. Start a table.
            // ------------------------------------------------------------
            builder.StartTable(); // start table

            // ------------------------------------------------------------
            // 4. Insert header row (column names).
            // ------------------------------------------------------------
            foreach (DataColumn col in data.Columns)
            {
                builder.InsertCell();               // insert a new cell
                builder.Write(col.ColumnName);      // write header text
            }
            builder.EndRow(); // finish header row

            // ------------------------------------------------------------
            // 5. Insert data rows.
            // ------------------------------------------------------------
            foreach (DataRow row in data.Rows)
            {
                foreach (object cellValue in row.ItemArray)
                {
                    builder.InsertCell();                     // insert a new cell
                    builder.Write(cellValue?.ToString() ?? ""); // write cell value
                }
                builder.EndRow(); // finish the current data row
            }

            // ------------------------------------------------------------
            // 6. End the table and save the document.
            // ------------------------------------------------------------
            builder.EndTable(); // end table
            doc.Save("TableFromDataTable.docx"); // save document
        }
    }
}
