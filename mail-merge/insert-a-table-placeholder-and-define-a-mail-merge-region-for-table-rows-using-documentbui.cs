using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeTableRegionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table that will contain the mail merge region.
            builder.StartTable();

            // Insert a single cell that will hold the region start, the merge fields, and the region end.
            builder.InsertCell();

            // Insert the TableStart field for the region named "Employees".
            builder.InsertField(" MERGEFIELD TableStart:Employees ");

            // Insert merge fields that correspond to the columns of the data source.
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write(" "); // Add a space between fields.
            builder.InsertField(" MERGEFIELD Age ");

            // Insert the TableEnd field to close the region.
            builder.InsertField(" MERGEFIELD TableEnd:Employees ");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Prepare a DataTable that matches the region name and contains data for the merge fields.
            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("Name", typeof(string));
            employees.Columns.Add("Age", typeof(int));

            // Add sample rows.
            employees.Rows.Add("John Doe", 30);
            employees.Rows.Add("Jane Smith", 25);
            employees.Rows.Add("Bob Johnson", 40);

            // Execute the mail merge with regions. The table will be expanded for each row in the DataTable.
            doc.MailMerge.ExecuteWithRegions(employees);

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeWithTableRegion.docx");
            doc.Save(outputPath);

            // Inform the user where the file was saved.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
