using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data in a DataTable.
        DataTable peopleTable = new DataTable("People");
        peopleTable.Columns.Add("Name", typeof(string));
        peopleTable.Columns.Add("Age", typeof(int));

        peopleTable.Rows.Add("Alice", 30);
        peopleTable.Rows.Add("Bob", 25);
        peopleTable.Rows.Add("Charlie", 35);

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a data band that iterates over the rows of the DataTable.
        builder.Writeln("<<foreach [row in dt]>>");
        // Output each column value for the current row.
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        // End the data band.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and generate the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // Use the LINQ Reporting engine to populate the template.
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("dt") must match the name used in the template tags.
        engine.BuildReport(report, peopleTable, "dt");

        // Save the generated report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}
