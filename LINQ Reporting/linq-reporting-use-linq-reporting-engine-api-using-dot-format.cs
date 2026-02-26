using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Define input template and output locations.
            string templatePath = @"C:\Templates\ReportTemplate.docx";
            string outputPath   = @"C:\Output\GeneratedReport.docx";
            string jsonDataPath = @"C:\Data\ReportData.json";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JSON data source from a file.
            JsonDataSource dataSource = new JsonDataSource(jsonDataPath);

            // Initialize the LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: configure engine options (e.g., remove empty paragraphs).
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The data source is referenced in the template as "data".
            engine.BuildReport(doc, dataSource, "data");

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
