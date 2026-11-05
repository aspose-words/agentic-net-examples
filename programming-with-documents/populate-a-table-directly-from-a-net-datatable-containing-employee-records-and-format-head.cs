using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace AsposeWordsTableFromDataTable
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample DataTable with employee records.
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("ID", typeof(int));
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Department", typeof(string));
            employeeTable.Columns.Add("Salary", typeof(decimal));

            employeeTable.Rows.Add(1, "Alice Johnson", "Finance", 75000m);
            employeeTable.Rows.Add(2, "Bob Smith", "IT", 68000m);
            employeeTable.Rows.Add(3, "Carol White", "HR", 62000m);
            employeeTable.Rows.Add(4, "David Brown", "Marketing", 71000m);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table into the document.
            Table table = builder.StartTable();

            // Define column widths (optional).
            builder.CellFormat.Width = 100; // ID column
            builder.InsertCell();
            builder.CellFormat.Width = 200; // Name column
            builder.InsertCell();
            builder.CellFormat.Width = 150; // Department column
            builder.InsertCell();
            builder.CellFormat.Width = 120; // Salary column
            builder.InsertCell();

            // Move back to the first cell to start adding header cells.
            builder.MoveTo(table.FirstRow.FirstCell);

            // Format header row: bold text, centered, light gray background.
            builder.Font.Bold = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

            // Header cells.
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
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Empty;

            // Populate table rows from the DataTable.
            foreach (DataRow row in employeeTable.Rows)
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
            string outputPath = "EmployeeReport.docx";
            doc.Save(outputPath);
        }
    }
}
