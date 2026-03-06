using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled template (DOTM) file.
            string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Path to the JSON file that will be used as the data source.
            string jsonDataPath = @"C:\Data\ReportData.json";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Output\GeneratedReport.dotm";

            // Load the DOTM template into a Document object.
            Document templateDoc = new Document(templatePath);

            // Create a JSON data source from the file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonDataPath);

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: set engine options, e.g., remove empty paragraphs after processing.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The second parameter is the data source,
            // the third parameter is the name used to reference the data source in the template.
            engine.BuildReport(templateDoc, jsonDataSource, "data");

            // Save the populated document back as a DOTM file.
            templateDoc.Save(outputPath, SaveFormat.Dotm);
        }
    }
}
