using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document (DOCX) that contains LINQ Reporting Engine tags.
            string templatePath = Path.Combine("Data", "ReportTemplate.docx");

            // Load the template document.
            Document doc = new Document(templatePath);

            // Path to the JSON file that will serve as the data source.
            string jsonPath = Path.Combine("Data", "People.json");

            // Create a JsonDataSource instance from the JSON file.
            JsonDataSource jsonData = new JsonDataSource(jsonPath);

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by merging the template with the JSON data.
            // The third argument ("persons") is the name used to reference the data source in the template.
            engine.BuildReport(doc, jsonData, "persons");

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Example option: export form fields as plain text.
                ExportTextInputFormFieldAsText = true,
                // Example option: specify a folder for extracted images.
                ImagesFolder = Path.Combine("Output", "Images")
            };

            // Ensure the images folder exists.
            Directory.CreateDirectory(htmlOptions.ImagesFolder);

            // Save the generated report as an HTML file.
            string outputPath = Path.Combine("Output", "Report.html");
            doc.Save(outputPath, htmlOptions);

            Console.WriteLine($"Report generated and saved to: {outputPath}");
        }
    }
}
