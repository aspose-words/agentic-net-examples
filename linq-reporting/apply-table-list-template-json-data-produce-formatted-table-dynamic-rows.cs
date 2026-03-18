using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsInTableListExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the template document and JSON data file.
            string templatePath = @"C:\Templates\InTableListTemplate.docx";
            string jsonPath = @"C:\Data\People.json";
            string outputPath = @"C:\Output\ReportWithDynamicRows.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JSON data source from the file.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report by populating the template with the JSON data.
            // The data source name "people" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "people");

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
