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
        DataTable table = new()
        {
            TableName = "Data"
        };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));

        table.Rows.Add(1, "Alice", 123.45m);
        table.Rows.Add(2, "Bob", 678.90m);
        table.Rows.Add(3, "Charlie", 234.56m);

        // Create a template document programmatically.
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a heading.
        builder.Writeln("Report generated from DataTable");
        builder.Writeln();

        // Insert a table header.
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();
        builder.EndTable();

        // Begin the foreach block that iterates over the DataTable rows.
        builder.Writeln("<<foreach [row in Data]>>");

        // Create a table for each row.
        Table dataTable = builder.StartTable();

        // Id cell.
        builder.InsertCell();
        builder.Writeln("<<[row.Id]>>");

        // Name cell.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");

        // Amount cell.
        builder.InsertCell();
        builder.Writeln("<<[row.Amount]>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report using the DataTable as the root data source.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, table, "Data");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
