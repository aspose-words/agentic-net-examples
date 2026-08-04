using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Enable code page provider for CSV parsing (required for non‑UTF8 encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create an output folder relative to the current directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file that will be used as the data source.
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,Email",
            "Alice,30,alice@example.com",
            "Bob,25,bob@example.com",
            "Charlie,35,charlie@example.com"
        });

        // -----------------------------------------------------------------
        // 2. Build a template document containing a foreach block and a table.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [person in persons]>>");

        // Table with placeholders for each CSV column.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Email");
        builder.EndRow();

        // Data row – will be repeated for each CSV record.
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Email]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Configure CSV data source (first line contains headers).
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None   // default options
        };
        engine.BuildReport(reportDoc, csvData, "persons");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "PeopleReport.docx");
        reportDoc.Save(outputPath);
    }
}
