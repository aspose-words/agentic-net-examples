using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains header/footer merge fields.
        // Example header/footer tags: <<[data.CompanyName]>>, <<[data.ReportDate]>>
        Document template = new Document("Template.docx");

        // Load JSON data. The JSON file should contain properties that match the merge fields.
        // For example: { "CompanyName": "Acme Corp", "ReportDate": "2024-12-01" }
        JsonDataSource jsonSource = new JsonDataSource("data.json");

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report, exposing the JSON data source under the name "data".
        // All merge fields in the main body, header, and footer that reference "data"
        // will be populated with the corresponding JSON values.
        engine.BuildReport(template, jsonSource, "data");

        // Save the populated document.
        template.Save("Report.docx");
    }
}
