using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonBindingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains Aspose.Words reporting tags.
            // Example tag in the template: <<[persons.Name]>>
            string templatePath = @"C:\Templates\CustomerReportTemplate.docx";

            // Path to the JSON file that holds the customer data.
            string jsonDataPath = @"C:\Data\Customers.json";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\CustomerReport.docx";

            // Load the template document.
            // This uses the Document(string) constructor – the approved "load" rule.
            Document template = new Document(templatePath);

            // Create a JSON data source from the file.
            // JsonDataSource(string) constructor follows the approved "create" rule.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonDataPath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by merging the JSON data into the template.
            // Use the overload that allows referencing the data source by name ("persons").
            // This follows the approved "build report" rule.
            engine.BuildReport(template, jsonDataSource, "persons");

            // Save the populated document.
            // Document.Save(string) follows the approved "save" rule.
            template.Save(outputPath);

            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
