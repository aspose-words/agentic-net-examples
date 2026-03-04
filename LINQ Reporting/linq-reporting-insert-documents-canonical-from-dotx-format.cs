using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple data class that will be used as the data source for the report.
    public class ReportItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOTX template that contains the reporting tags.
            // The template can reference the data source as <<[src.Id]>> and <<[src.Name]>>.
            Document template = new Document("Template.dotx");

            // Create a collection of data items using LINQ.
            List<ReportItem> dataSource = Enumerable.Range(1, 5)
                .Select(i => new ReportItem
                {
                    Id = i,
                    Name = $"Item {i}"
                })
                .ToList();

            // Build the report. The data source name "src" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "src");

            // Save the populated document.
            template.Save("ReportResult.docx");
        }
    }
}
