using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains LINQ Reporting Engine tags.
            // Example tag in the template: <<[persons.Name]>>
            string templatePath = @"Template.docx";

            // Path to the JSON data source file.
            // Example JSON content:
            // [
            //   { "Name": "John Doe", "Age": 30 },
            //   { "Name": "Jane Smith", "Age": 25 }
            // ]
            string jsonDataPath = @"people.json";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JSON data source from the file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonDataPath);

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The data source name "persons" is used in the template tags.
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the generated report as DOCX.
            string outputPath = @"Report.docx";
            doc.Save(outputPath);
        }
    }
}
