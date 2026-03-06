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
            // Sample DataTable creation
            DataTable dt = new DataTable("Sample");
            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("Country");
            dt.Rows.Add("Alice", 30, "USA");
            dt.Rows.Add("Bob", 25, "UK");
            dt.Rows.Add("Charlie", 35, "Canada");

            // Create a new Word document and fill it with a table based on the DataTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table
            Table table = builder.StartTable();

            // Optional: add a header row
            foreach (DataColumn col in dt.Columns)
            {
                builder.InsertCell();
                builder.Write(col.ColumnName);
            }
            builder.EndRow();

            // Iterate through each DataRow and add a new table row
            foreach (DataRow row in dt.Rows)
            {
                foreach (object cellValue in row.ItemArray)
                {
                    builder.InsertCell();
                    // Convert the cell value to string, handling nulls
                    builder.Write(cellValue?.ToString() ?? string.Empty);
                }
                builder.EndRow();
            }

            // Finish the table
            builder.EndTable();

            // Save the document (replace with your desired path)
            doc.Save("TableFromDataTable.docx");
        }
    }
}
