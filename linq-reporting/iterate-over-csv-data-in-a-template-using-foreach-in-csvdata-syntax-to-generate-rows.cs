using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple CSV file with headers.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath,
            "Name,Age,Country\n" +
            "Alice,30,USA\n" +
            "Bob,25,Canada\n" +
            "Charlie,35,UK");

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a foreach tag that iterates over the CSV data source named "csvData".
        builder.Writeln("<<foreach [in csvData]>>");
        // Inside the loop output each column value.
        builder.Writeln("<<[Name]>>\t<<[Age]>>\t<<[Country]>>");
        builder.Writeln("<</foreach>>");

        // Load the CSV data as a data source.
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, csvDataSource, "csvData");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}
