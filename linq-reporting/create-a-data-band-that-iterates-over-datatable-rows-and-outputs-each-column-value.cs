using System;
using System.Data;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample DataTable
        DataTable table = new DataTable("Sample");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Value", typeof(double));

        table.Rows.Add(1, "Alpha", 12.34);
        table.Rows.Add(2, "Beta", 56.78);
        table.Rows.Add(3, "Gamma", 90.12);

        // 2. Build the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Data Table Report");
        builder.Writeln(); // empty line

        // Start data band that iterates over DataTable rows
        builder.Writeln("<<foreach [row in table]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport)
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // 3. Load the template and generate the report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the DataTable as the root data source
        engine.BuildReport(reportDoc, table, "table");

        // 4. Save the generated report
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
