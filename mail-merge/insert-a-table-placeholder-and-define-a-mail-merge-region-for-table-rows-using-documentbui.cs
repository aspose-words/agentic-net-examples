using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace MailMergeTableRegionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table that will contain the mail merge region.
            builder.StartTable();

            // Insert the TableStart field for the region named "Employees".
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableStart:Employees ");

            // Insert a cell for the "Name" column.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Name ");

            // Insert a cell for the "Age" column.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Age ");

            // Insert the TableEnd field to close the region.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableEnd:Employees ");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Prepare a DataTable that matches the region name and column names.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("Name", typeof(string));
            employees.Columns.Add("Age", typeof(int));

            employees.Rows.Add("John Doe", 30);
            employees.Rows.Add("Jane Smith", 27);
            employees.Rows.Add("Bob Johnson", 45);

            // Execute the mail merge with regions using the DataTable.
            doc.MailMerge.ExecuteWithRegions(employees);

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeTableRegion.docx");
            doc.Save(outputPath);
        }
    }
}
