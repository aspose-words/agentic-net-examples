using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create a sample CSV file.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Quantity",
            "1,Apple,10",
            "2,Banana,20",
            "3,Cherry,15"
        });

        // 2. Build a Word template containing a table with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln(); // blank line

        // Begin the foreach block that iterates over the CSV rows.
        builder.Writeln("<<foreach [row in data]>>");

        // Create the table.
        Table table = builder.StartTable();

        // Header row (static content).
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.EndRow();

        // Data row (dynamic content).
        builder.InsertCell(); builder.Writeln("<<[row.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Quantity]>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template and bind the CSV data source.
        var reportDoc = new Document(templatePath);

        var loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ','
        };

        var csvData = new CsvDataSource(csvPath, loadOptions);

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // 4. Save the generated report.
        reportDoc.Save("output.docx");
    }
}
