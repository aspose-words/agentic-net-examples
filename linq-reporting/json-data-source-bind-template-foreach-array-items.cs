using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonReport
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains the <<foreach>> tags.
            // Example template content:
            //   <<foreach [persons]>>
            //       <<[FirstName]>> <<[LastName]>>
            //   <</foreach>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path to the JSON file that holds the data.
            // The JSON should have a top‑level array named "persons":
            //   {
            //       "persons": [
            //           { "FirstName": "John", "LastName": "Doe" },
            //           { "FirstName": "Jane", "LastName": "Smith" }
            //       ]
            //   }
            string jsonPath = @"C:\Data\People.json";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Optional: configure JSON parsing options.
            JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
            {
                // Preserve spaces in string values if needed.
                PreserveSpaces = false,
                // Keep the root object even if the JSON root is a single array.
                AlwaysGenerateRootObject = true,
                // Use loose parsing so numbers in strings are converted to numeric types.
                SimpleValueParseMode = JsonSimpleValueParseMode.Loose
            };

            // Create a JSON data source using the file path and the options above.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

            // Build the report. The third argument is the name used in the template to reference the data source.
            // In the template we used <<foreach [persons]>> so we pass "persons" as the data source name.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the generated report.
            string outputPath = @"C:\Output\ReportResult.docx";
            doc.Save(outputPath);
        }
    }
}
