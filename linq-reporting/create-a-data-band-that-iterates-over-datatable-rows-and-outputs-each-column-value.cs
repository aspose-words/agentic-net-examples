using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataTable.
        DataTable dataTable = new DataTable("People");
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Rows.Add("Alice", 30);
        dataTable.Rows.Add("Bob", 25);
        dataTable.Rows.Add("Charlie", 35);

        // Create a Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting data band that iterates over the rows of the DataTable.
        builder.Writeln("<<foreach [row in dt]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the saved template.
        Document report = new Document(templatePath);

        // Build the report using the DataTable as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, dataTable, "dt");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
