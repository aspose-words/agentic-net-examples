using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class CsvReportExample
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a sample CSV file.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Quantity,Price",
            "1,Apple,10,0.5",
            "2,Banana,5,0.3",
            "3,Carrot,7,0.2"
        });

        // 2. Build the template document with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Report generated from CSV data");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [row in data]>>");

        // Start table.
        Table table = builder.StartTable();

        // Insert a temporary row so that the table is not empty.
        builder.InsertCell();
        builder.Writeln(string.Empty);
        builder.EndRow();

        // Now the table has at least one row – apply the style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Header row (will appear once, outside the data rows).
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Price");
        builder.EndRow();

        // Data row – each cell contains a tag that reads a field from the current CSV record.
        builder.InsertCell(); builder.Writeln("<<[row.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Price]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // 3. Load the template and bind the CSV data source.
        Document report = new Document(templatePath);

        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            HasHeaders = true
        };
        CsvDataSource csvSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, csvSource, "data");

        // 4. Save the final report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        report.Save(reportPath);
    }
}
