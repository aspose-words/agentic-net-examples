using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains a table with reporting tags.
        // The template should define a repeatable region that iterates over the JSON data,
        // e.g. <<foreach [data]>> for rows and <<[ColumnNames]>> for header cells.
        Document template = new Document("Template.docx");

        // Create a JSON data source from a file.
        // The constructor JsonDataSource(string) uses default parsing options.
        JsonDataSource jsonSource = new JsonDataSource("data.json");

        // Build the report by merging the JSON data into the template.
        // The third argument is the name used inside the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, jsonSource, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
