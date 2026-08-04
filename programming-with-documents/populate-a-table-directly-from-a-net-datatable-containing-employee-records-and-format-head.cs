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
            // Create a DataTable with employee data.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("ID", typeof(int));
            employees.Columns.Add("Name", typeof(string));
            employees.Columns.Add("Position", typeof(string));
            employees.Columns.Add("Salary", typeof(decimal));

            employees.Rows.Add(1, "John Doe", "Developer", 75000);
            employees.Rows.Add(2, "Jane Smith", "Designer", 68000);
            employees.Rows.Add(3, "Bob Johnson", "Manager", 82000);
            employees.Rows.Add(4, "Alice Brown", "Tester", 59000);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Format the header row.
            builder.RowFormat.Height = 20;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;
            builder.Font.Bold = true;

            // Insert header cells.
            InsertHeaderCell(builder, "ID");
            InsertHeaderCell(builder, "Name");
            InsertHeaderCell(builder, "Position");
            InsertHeaderCell(builder, "Salary");
            builder.EndRow();

            // Reset formatting for data rows.
            builder.Font.Bold = false;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;

            // Populate the table with DataTable rows.
            foreach (DataRow row in employees.Rows)
            {
                builder.InsertCell();
                builder.Write(row["ID"].ToString());

                builder.InsertCell();
                builder.Write(row["Name"].ToString());

                builder.InsertCell();
                builder.Write(row["Position"].ToString());

                builder.InsertCell();
                builder.Write(string.Format("{0:C}", row["Salary"]));

                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("EmployeeTable.docx");
        }

        // Helper method to insert a header cell with consistent formatting.
        private static void InsertHeaderCell(DocumentBuilder builder, string text)
        {
            builder.InsertCell();
            builder.Write(text);
        }
    }
}
