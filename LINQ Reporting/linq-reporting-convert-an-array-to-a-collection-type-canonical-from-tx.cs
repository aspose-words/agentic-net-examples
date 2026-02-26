using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains reporting tags.
            string templatePath = @"C:\Docs\ReportTemplate.docx";

            // Path to the plain‑text file (TXT) that holds the raw data, one value per line.
            string txtDataPath = @"C:\Docs\SourceData.txt";

            // Load the template document.
            Document template = new Document(templatePath);

            // Read all lines from the TXT file into a string array.
            string[] rawLines = File.ReadAllLines(txtDataPath);

            // Convert the array into a collection that the LINQ Reporting Engine can work with.
            // Here we create a list of anonymous objects with two fields: Index and Value.
            // The fields can be referenced in the template as <<[data.Index]>> and <<[data.Value]>>.
            var dataCollection = rawLines
                .Select((line, idx) => new { Index = idx + 1, Value = line })
                .ToList();

            // Initialise the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the collection as a data source.
            // The second parameter is the name by which the data source will be referenced in the template.
            engine.BuildReport(template, dataCollection, "data");

            // Save the generated report.
            string outputPath = @"C:\Docs\GeneratedReport.docx";
            template.Save(outputPath);
        }
    }
}
