using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataTable.
        DataTable data = new DataTable("Data");
        data.Columns.Add("FirstName", typeof(string));
        data.Columns.Add("LastName", typeof(string));
        data.Columns.Add("Age", typeof(int));

        data.Rows.Add("John", "Doe", 30);
        data.Rows.Add("Jane", "Smith", 28);
        data.Rows.Add("Bob", "Johnson", 45);

        // ---------------------------------------------------------------
        // 1. Create a template document that contains LINQ Reporting tags.
        // ---------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the rows of the DataTable.
        builder.Writeln("<<foreach [row in Data]>>");

        // Build a simple table where each cell contains a LINQ expression that reads a column value.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("First Name");
        builder.InsertCell();
        builder.Writeln("Last Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[row.FirstName]>>");

        builder.InsertCell();
        builder.Writeln("<<[row.LastName]>>");

        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");

        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // ---------------------------------------------------------------
        // 2. Load the template and build the report using ReportingEngine.
        // ---------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The data source name used in the template tags is "Data".
        engine.BuildReport(report, data, "Data");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
