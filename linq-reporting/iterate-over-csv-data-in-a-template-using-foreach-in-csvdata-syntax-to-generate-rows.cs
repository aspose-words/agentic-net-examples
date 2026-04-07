using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type
using System.Text;          // For Encoding registration

public class CsvLinqReportingExample
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Age",
            "1,John Doe,30",
            "2,Jane Smith,25",
            "3,Bob Johnson,40"
        });

        // Create a Word template programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Static table header (not repeated).
        Table headerTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Age");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln();

        // Begin foreach loop over the CSV data source.
        // Correct syntax: <<foreach [item in csvData]>>
        builder.Writeln("<<foreach [person in csvData]>>");

        // Table row that will be repeated for each CSV record.
        Table dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[person.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[person.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[person.Age]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Load CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "csvData");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PeopleReport.docx");
        doc.Save(outputPath);
    }
}
