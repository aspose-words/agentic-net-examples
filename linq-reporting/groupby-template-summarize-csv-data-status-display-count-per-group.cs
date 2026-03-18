using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that uses the Reporting Engine syntax.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Summary by Status");
        builder.Writeln();

        // GroupBy block – groups rows by the \"Status\" column and outputs the count per group.
        // The data source will be referenced as \"persons\" (the name we will use in BuildReport).
        builder.Writeln("<<groupby [persons] Status>>");
        builder.Writeln("Status: <<[Status]>> – Count: <<count>>");
        builder.Writeln("<<endgroupby>>");

        // Load CSV data. The CSV file must have a header row (HasHeaders = true).
        // Example CSV content:
        // Status,Name
        // Open,John
        // Closed,Anna
        // Open,Mark
        var loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ','; // default, set explicitly for clarity
        CsvDataSource csvSource = new CsvDataSource("data.csv", loadOptions);

        // Build the report using the template and the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, csvSource, "persons");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
