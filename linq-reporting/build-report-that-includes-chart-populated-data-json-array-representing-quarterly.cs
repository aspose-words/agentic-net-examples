using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportGenerator
{
    static void Main()
    {
        // Path to the Word template that contains a chart with template tags.
        const string templatePath = "Template.docx";

        // Path to the JSON file that holds the quarterly results.
        const string jsonPath = "QuarterlyResults.json";

        // Path where the generated report will be saved.
        const string outputPath = "ReportWithChart.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a JSON data source from the file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report by merging the JSON data into the template.
        // The data source is referenced in the template with the name "results".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "results");

        // Save the populated document.
        doc.Save(outputPath);
    }
}
