using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a sample CSV file with a header and some rows.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Status",
            "1,Apple,Available",
            "2,Banana,OutOfStock",
            "3,Cherry,Available",
            "4,Date,Available",
            "5,Elderberry,OutOfStock"
        });

        // Build a template document that groups rows by Status and shows the count per group.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Status Summary:");
        builder.Writeln("<<foreach [g in persons.GroupBy(p => p.Status)]>>");
        builder.Writeln("Status: <<[g.Key]>> - Count: <<[g.Count()]>>");
        builder.Writeln("<</foreach>>");

        // Load the CSV data source, indicating that the first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Generate the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "persons");

        // Save the populated document.
        template.Save("Report.docx");
    }
}
