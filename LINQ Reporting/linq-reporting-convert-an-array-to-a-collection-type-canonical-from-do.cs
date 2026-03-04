using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple data model used as the data source for the LINQ Reporting Engine.
    public class ReportData
    {
        // Original array of values.
        public int[] ValuesArray { get; set; }

        // Canonical collection type required by the reporting engine (IEnumerable<int>).
        // The getter converts the array to a List<int> on demand.
        public IEnumerable<int> ValuesCollection => ValuesArray?.ToList() ?? Enumerable.Empty<int>();
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains LINQ Reporting tags.
            Document template = new Document("Template.dotm");

            // Prepare the data source with an array of integers.
            var data = new ReportData
            {
                ValuesArray = new[] { 10, 20, 30, 40, 50 }
            };

            // Create the reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the template and the data source.
            // The template can reference the collection via the property "ValuesCollection".
            engine.BuildReport(template, data, "ds");

            // Save the generated report.
            template.Save("ReportOutput.docx");
        }
    }
}
