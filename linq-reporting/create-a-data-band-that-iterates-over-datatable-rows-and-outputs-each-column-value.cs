using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        DataTable data = new();
        data.Columns.Add("Id", typeof(int));
        data.Columns.Add("Name", typeof(string));
        data.Columns.Add("Value", typeof(double));

        data.Rows.Add(1, "Alpha", 12.34);
        data.Rows.Add(2, "Beta", 56.78);
        data.Rows.Add(3, "Gamma", 90.12);

        // Create the LINQ Reporting template.
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Title.
        builder.Writeln("Data Table Report");
        builder.Writeln();

        // Begin the foreach data band – the whole table will be repeated for each row.
        builder.Writeln("<<foreach [row in data]>>");

        // Start the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row – values are taken from the current DataRow.
        builder.InsertCell();
        builder.Writeln("<<[row.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Value]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report using the DataTable as the root data source.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, data, "data");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
