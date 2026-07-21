using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataTable.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName", typeof(string));
        table.Columns.Add("LastName", typeof(string));
        table.Columns.Add("Age", typeof(int));

        table.Rows.Add("John", "Doe", 30);
        table.Rows.Add("Jane", "Smith", 28);
        table.Rows.Add("Bob", "Johnson", 45);

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // The template contains a foreach loop that iterates over the rows
        // of the DataTable (named "dt" in the report) and writes three
        // fields for FirstName, LastName and Age.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Employee List:");
        // Begin foreach block.
        builder.Writeln("<<foreach [row in dt]>>");

        // Write the fields for each row.
        builder.Write("<<[row.FirstName]>> ");
        builder.Write("<<[row.LastName]>>");
        builder.Write(", Age: ");
        builder.Write("<<[row.Age]>>");

        // End the paragraph for the current row.
        builder.Writeln();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The data source name used in the template tags is "dt".
        engine.BuildReport(report, table, "dt");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
