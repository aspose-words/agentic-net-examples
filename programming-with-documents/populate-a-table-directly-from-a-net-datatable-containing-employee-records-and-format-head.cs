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
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("ID", typeof(int));
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Department", typeof(string));
            employeeTable.Columns.Add("Salary", typeof(decimal));

            // Add sample rows.
            employeeTable.Rows.Add(1, "Alice Johnson", "Finance", 75000m);
            employeeTable.Rows.Add(2, "Bob Smith", "IT", 68000m);
            employeeTable.Rows.Add(3, "Carol White", "HR", 62000m);
            employeeTable.Rows.Add(4, "David Brown", "Marketing", 71000m);

            // Create a new Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            builder.StartTable();

            // Build header row.
            foreach (DataColumn column in employeeTable.Columns)
            {
                builder.InsertCell();
                builder.Write(column.ColumnName);
            }
            builder.EndRow();

            // Populate table rows from the DataTable.
            foreach (DataRow row in employeeTable.Rows)
            {
                foreach (object cellValue in row.ItemArray)
                {
                    builder.InsertCell();
                    builder.Write(cellValue?.ToString() ?? string.Empty);
                }
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Retrieve the created table from the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Apply formatting to the header row.
            Row headerRow = table.FirstRow;
            foreach (Cell cell in headerRow.Cells)
            {
                // Make header text bold.
                if (cell.FirstParagraph?.Runs.Count > 0)
                {
                    cell.FirstParagraph.Runs[0].Font.Bold = true;
                }

                // Set background shading for header cells.
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                // Optional: set vertical alignment.
                cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            }

            // Apply a simple border to the whole table.
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
            table.SetBorder(BorderType.Horizontal, LineStyle.Single, 0.5, Color.Gray, true);
            table.SetBorder(BorderType.Vertical, LineStyle.Single, 0.5, Color.Gray, true);

            // Save the document.
            string outputPath = "EmployeeReport.docx";
            doc.Save(outputPath);
        }
    }
}
