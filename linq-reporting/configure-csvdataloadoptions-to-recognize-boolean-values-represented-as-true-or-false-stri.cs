using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing on .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with boolean values represented as "true"/"false".
        string csvPath = "people.csv";
        File.WriteAllText(csvPath,
            "Name,IsActive\n" +
            "Alice,true\n" +
            "Bob,false\n" +
            "Charlie,true");

        // Configure CSV loading options.
        // - HasHeaders = true tells the engine that the first line contains column names.
        // - Delimiter = ',' (default) separates fields.
        // - QuoteChar = '\"' (default) handles quoted values.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '\"',
            // No comment character needed for this simple file.
        };

        // Create a CSV data source using the configured options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("People Report");
        builder.Writeln("----------------");

        // Iterate over the rows of the CSV data source.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Active: <<[p.IsActive]>>"); // Boolean values will be output as True/False.
        builder.Writeln("<</foreach>>");

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        // The data source name "persons" matches the name used in the template tags.
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("PeopleReport.docx");
    }
}
