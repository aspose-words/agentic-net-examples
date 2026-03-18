using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonReportGenerator
{
    static void Main()
    {
        // Paths to the template, JSON data file and the output document.
        string templatePath = @"C:\Docs\ReportTemplate.docx";
        string jsonPath = @"C:\Docs\PeopleData.json";
        string outputPath = @"C:\Docs\GeneratedReport.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure JSON loading options to enable type inference (Loose mode).
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose
        };

        // Create a JsonDataSource from the JSON file using the options above.
        JsonDataSource jsonData = new JsonDataSource(jsonPath, loadOptions);

        // Build the report. The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
