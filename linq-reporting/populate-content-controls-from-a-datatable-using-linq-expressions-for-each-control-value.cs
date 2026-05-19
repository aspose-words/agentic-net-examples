using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -------------------------
        // 1. Create the template document.
        // -------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header.
        builder.Writeln("Report generated from DataTable:");
        builder.Writeln();

        // Begin foreach over the DataTable rows (named "Data").
        builder.Writeln("<<foreach [row in Data]>>");

        // Row content: Title - Description: Amount
        builder.Writeln("<<[row.Title]>> - <<[row.Description]>>: <<[row.Amount]>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template locally.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -------------------------
        // 2. Prepare sample data in a DataTable.
        // -------------------------
        DataTable data = new();
        data.Columns.Add("Title", typeof(string));
        data.Columns.Add("Description", typeof(string));
        data.Columns.Add("Amount", typeof(decimal));

        data.Rows.Add("Item A", "First item description", 123.45m);
        data.Rows.Add("Item B", "Second item description", 678.90m);
        data.Rows.Add("Item C", "Third item description", 0m);

        // -------------------------
        // 3. Load the template and build the report using LINQ Reporting.
        // -------------------------
        Document report = new(templatePath);

        ReportingEngine engine = new();
        // BuildReport overload that takes a data source name.
        engine.BuildReport(report, data, "Data");

        // -------------------------
        // 4. Save the final document.
        // -------------------------
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
