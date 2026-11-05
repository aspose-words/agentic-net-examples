using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create sample DataTable.
        DataTable table = new DataTable("People");
        table.Columns.Add("FirstName", typeof(string));
        table.Columns.Add("LastName", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("John", "Doe", 30);
        table.Rows.Add("Jane", "Smith", 25);
        table.Rows.Add("Bob", "Johnson", 40);

        // Build the LINQ Reporting template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over the DataTable rows.
        builder.Writeln("<<foreach [row in Data]>>");

        // Create a table with a header row.
        Table tbl = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("First Name");
        builder.InsertCell();
        builder.Writeln("Last Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row with LINQ Reporting tags bound to each column.
        builder.InsertCell();
        builder.Writeln("<<[row.FirstName]>>");

        builder.InsertCell();
        builder.Writeln("<<[row.LastName]>>");

        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and generate the report.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, table, "Data");

        // Save the final document.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
