using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the Word template that contains the single‑row markup.
        // Example template content:
        // <<foreach [persons]>>
        //   Name: <<[Name]>>
        //   Age: <<[Age]>>
        // <<endforeach>>
        string templatePath = @"C:\Templates\PersonReport.docx";

        // Path to the JSON file that holds an array of person objects.
        // Example JSON:
        // [
        //   { "Name": "John Doe", "Age": 30 },
        //   { "Name": "Jane Smith", "Age": 25 }
        // ]
        string jsonPath = @"C:\Data\people.json";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a JsonDataSource from the JSON file using default parsing options.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report. The data source name ("persons") must match the name used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "persons");

        // Save the generated report.
        string outputPath = @"C:\Output\PersonReportGenerated.docx";
        doc.Save(outputPath);
    }
}
