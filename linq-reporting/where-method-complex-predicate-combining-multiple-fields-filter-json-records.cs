// Load a Word template that contains a Reporting Engine expression using LINQ Where.
// Example template syntax (inside the document):
//   {{persons.Where(p => p.Age > 30 && p.City == "London")}}
//
// The code below prepares the JSON data source, configures parsing options,
// and builds the report. The filtering is performed entirely by the
// template expression; the C# code only supplies the raw JSON data.

using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonWhereExample
{
    static void Main()
    {
        // Path to the Word template that contains the {{persons.Where(...)}}
        // expression. The template must have a mail‑merge region or a table that
        // iterates over the collection named "persons".
        string templatePath = @"C:\Docs\ReportTemplate.docx";

        // Path to the JSON file that holds an array of person objects.
        // Example JSON:
        // [
        //   { "Name": "John", "Age": 45, "City": "London" },
        //   { "Name": "Anna", "Age": 28, "City": "Paris" },
        //   ...
        // ]
        string jsonPath = @"C:\Data\people.json";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure JSON parsing options if needed (e.g., preserve spaces,
        // custom date formats, or loose/simple value parsing).
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            // Loose mode interprets numeric strings as numbers, which is
            // convenient for numeric comparisons inside the Where predicate.
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose,
            PreserveSpaces = true
        };

        // Create the JSON data source. The root object name used in the template
        // ("persons") must match the name supplied to BuildReport.
        JsonDataSource jsonData = new JsonDataSource(jsonPath, loadOptions);

        // Build the report. The ReportingEngine evaluates the Where clause at
        // runtime, selecting only those records where Age > 30 AND City == "London".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "persons");

        // Save the generated document.
        string outputPath = @"C:\Docs\FilteredReport.docx";
        doc.Save(outputPath);
    }
}
