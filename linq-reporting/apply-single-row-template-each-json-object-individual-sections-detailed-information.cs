using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the Word template that contains a single‑row section.
        // The template should use Reporting Engine syntax, e.g. <<[people.Name]>>, <<[people.Age]>> etc.
        string templatePath = "Template.docx";

        // Path to the JSON file that holds an array of objects.
        // Example: [{ "Name":"John","Age":30 },{ "Name":"Jane","Age":25 }]
        string jsonPath = "People.json";

        // Path where the final report will be saved.
        string outputPath = "Report.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure JSON loading options.
        // AlwaysGenerateRootObject = true makes the engine treat the top‑level array as a table,
        // so each JSON object becomes a separate row (and thus a separate section in the template).
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };

        // Create a JSON data source using the file and the load options.
        JsonDataSource jsonData = new JsonDataSource(jsonPath, loadOptions);

        // Build the report. The data source name ("people") must match the name used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "people");

        // Save the populated document.
        doc.Save(outputPath);
    }
}
