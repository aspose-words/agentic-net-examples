using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportDemo
{
    class Program
    {
        static void Main()
        {
            // Load the Word template that contains the reporting tags.
            // Example tag in the template: <<[ds.TotalAmount]:$#,##0.00>>
            Document doc = new Document("Template.docx");

            // Load JSON data source. The JSON file should contain an object named "ds"
            // with a property "TotalAmount" (or any other numeric fields you need).
            JsonDataSource jsonData = new JsonDataSource("Data.json");

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after merging.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(doc, jsonData, "ds");

            // Save the populated document as DOCX.
            doc.Save("Report.docx");
        }
    }
}
