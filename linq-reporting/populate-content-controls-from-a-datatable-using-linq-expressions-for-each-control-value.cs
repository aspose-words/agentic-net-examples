using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace AsposeWordsLinqReportingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Customer Report");
            builder.Writeln();

            // LINQ Reporting foreach tag – iterate over rows of the DataTable named "Data".
            builder.Writeln("<<foreach [row in Data]>>");

            // Build a simple table with headers.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("First Name");
            builder.InsertCell();
            builder.Writeln("Last Name");
            builder.InsertCell();
            builder.Writeln("Age");
            builder.EndRow();

            // Data row – each cell contains a LINQ Reporting expression.
            builder.InsertCell();
            builder.Writeln("<<[row.FirstName]>>");
            builder.InsertCell();
            builder.Writeln("<<[row.LastName]>>");
            builder.InsertCell();
            builder.Writeln("<<[row.Age]>>");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Close the foreach block.
            builder.Writeln("<</foreach>>");

            // -----------------------------------------------------------------
            // Prepare sample data in a DataTable.
            DataTable dataTable = new DataTable("Data");
            dataTable.Columns.Add("FirstName", typeof(string));
            dataTable.Columns.Add("LastName", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));

            dataTable.Rows.Add("John", "Doe", 30);
            dataTable.Rows.Add("Jane", "Smith", 25);
            dataTable.Rows.Add("Bob", "Johnson", 40);

            // -----------------------------------------------------------------
            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options

            // The root object name ("Data") must match the name used in the template tags.
            engine.BuildReport(doc, dataTable, "Data");

            // Save the generated report.
            doc.Save("CustomerReport.docx");
        }
    }
}
