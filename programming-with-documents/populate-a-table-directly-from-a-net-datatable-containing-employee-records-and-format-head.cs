using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
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

            // Start the table.
            Table table = builder.StartTable();

            // Format the header row.
            builder.Font.Bold = true;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.InsertCell();
            builder.Write("ID");
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Position");
            builder.InsertCell();
            builder.Write("Salary");
            builder.EndRow();

            // Reset formatting for data rows.
            builder.Font.Bold = false;
            builder.CellFormat.Shading.ClearFormatting();

            // Populate the table with data from the DataTable.
            foreach (DataRow row in employees.Rows)
            {
                builder.InsertCell();
                builder.Write(row["ID"].ToString());

                builder.InsertCell();
                builder.Write(row["Name"].ToString());

                builder.InsertCell();
                builder.Write(row["Position"].ToString());

                builder.InsertCell();
                builder.Write(row["Salary"].ToString());

                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document.
            doc.Save("EmployeeTable.docx");
        }
    }
}
