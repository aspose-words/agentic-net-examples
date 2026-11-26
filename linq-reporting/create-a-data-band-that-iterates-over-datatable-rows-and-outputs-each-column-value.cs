using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data in a DataTable.
        DataTable dataTable = new DataTable("Sample");
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Value", typeof(double));

        dataTable.Rows.Add(1, "Alpha", 12.34);
        dataTable.Rows.Add(2, "Beta", 56.78);
        dataTable.Rows.Add(3, "Gamma", 90.12);

        // Create a LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Data Report");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the DataTable as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, dataTable, "data");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
