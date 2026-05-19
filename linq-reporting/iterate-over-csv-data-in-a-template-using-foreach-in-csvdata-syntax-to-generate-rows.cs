using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code page support for CSV parsing on .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a simple CSV file that will serve as the data source.
        // -----------------------------------------------------------------
        string csvFile = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        File.WriteAllLines(csvFile, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        });

        // ---------------------------------------------------------------
        // 2. Build a Word template in memory that contains LINQ Reporting
        //    tags. The foreach tag iterates over the CSV data source.
        // ---------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [in csvData]>>");   // Start loop over CSV rows
        builder.Writeln("<<[Name]>> - <<[Age]>>");    // Output fields from each row
        builder.Writeln("<</foreach>>");              // End loop

        // ---------------------------------------------------------------
        // 3. Load the CSV data using Aspose.Words.Reporting.CsvDataSource.
        // ---------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        CsvDataSource dataSource = new CsvDataSource(csvFile, loadOptions);

        // ---------------------------------------------------------------
        // 4. Execute the report. The data source name must match the name
        //    used in the template tags ("csvData").
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "csvData");

        // ---------------------------------------------------------------
        // 5. Save the generated report.
        // ---------------------------------------------------------------
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(outputFile);
    }
}
