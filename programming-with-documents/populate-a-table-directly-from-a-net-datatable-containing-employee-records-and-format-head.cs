using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFromDataTable
{
    public class Program
    {
        public static void Main()
        {
            // Create a DataTable with employee data.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("ID", typeof(int));
            employees.Columns.Add("Name", typeof(string));
            employees.Columns.Add("Department", typeof(string));
            employees.Columns.Add("Salary", typeof(decimal));

            employees.Rows.Add(1, "John Smith", "Finance", 75000);
            employees.Rows.Add(2, "Emily Davis", "HR", 68000);
            employees.Rows.Add(3, "Michael Brown", "IT", 82000);
            employees.Rows.Add(4, "Sarah Wilson", "Marketing", 71000);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // ---------- Header Row ----------
            // Apply header formatting: bold text and light gray background.
            builder.Font.Bold = true;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

            // Insert header cells.
            builder.InsertCell();
            builder.Write("ID");
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Department");
            builder.InsertCell();
            builder.Write("Salary");
            builder.EndRow();

            // Reset formatting for data rows.
            builder.Font.Bold = false;
            builder.CellFormat.Shading.ClearFormatting();

            // ---------- Data Rows ----------
            foreach (DataRow row in employees.Rows)
            {
                builder.InsertCell();
                builder.Write(row["ID"].ToString());

                builder.InsertCell();
                builder.Write(row["Name"].ToString());

                builder.InsertCell();
                builder.Write(row["Department"].ToString());

                builder.InsertCell();
                builder.Write(string.Format("{0:C}", row["Salary"]));

                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("EmployeeTable.docx");
        }
    }
}
