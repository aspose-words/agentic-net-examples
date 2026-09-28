using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data in a DataTable
        DataTable dataTable = new()
        {
            TableName = "Data"
        };
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Value", typeof(double));

        dataTable.Rows.Add(1, "Alpha", 12.34);
        dataTable.Rows.Add(2, "Beta", 56.78);
        dataTable.Rows.Add(3, "Gamma", 90.12);

        // Create the LINQ Reporting template programmatically
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin a data band that iterates over DataTable rows
        builder.Writeln("<<foreach [row in Data]>>");
        // Output each column value for the current row
        builder.Writeln("Id: <<[row.Id]>>\tName: <<[row.Name]>>\tValue: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation
        Document reportDoc = new(templatePath);

        // Build the report using the DataTable as the root data source
        ReportingEngine engine = new();
        bool success = engine.BuildReport(reportDoc, dataTable, "Data");

        // Save the generated report
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Optionally indicate success (no interactive prompts)
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
