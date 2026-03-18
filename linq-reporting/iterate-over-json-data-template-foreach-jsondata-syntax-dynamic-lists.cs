using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonForeachDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains the <<foreach>> tag.
            // Example template content (in the document body):
            //   <<foreach [in persons]>>
            //   Name: <<[Name]>>
            //   Age: <<[Age]>>
            //   <</foreach>>
            string templatePath = @"C:\Templates\PersonsTemplate.docx";

            // Path to the JSON file that will be used as the data source.
            // Example JSON:
            // {
            //   "persons": [
            //     { "Name": "John Doe", "Age": 30 },
            //     { "Name": "Jane Smith", "Age": 25 }
            //   ]
            // }
            string jsonPath = @"C:\Data\persons.json";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JsonDataSource from the JSON file using default parsing options.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report. The third argument ("persons") is the name that will be
            // referenced inside the template (the collection name used in <<foreach>>).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "persons");

            // Save the populated document.
            string outputPath = @"C:\Output\PersonsReport.docx";
            doc.Save(outputPath);
        }
    }
}
