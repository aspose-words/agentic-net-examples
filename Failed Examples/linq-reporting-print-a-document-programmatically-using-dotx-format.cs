// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsPrintExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template that will be used as the report source.
            // The Document constructor handles loading; no custom loading code is required.
            Document report = new Document("Template.dotx");

            // Prepare a simple data source for the LINQ reporting engine.
            // Here we use an anonymous object; any supported data source type can be used.
            var dataSource = new
            {
                Title = "Quarterly Sales Report",
                Date = DateTime.Now,
                Total = 123456.78
            };

            // Populate the template with data using the ReportingEngine.
            // BuildReport follows the provided API contract.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, dataSource, "ds"); // "ds" is the name used in the template.

            // Ensure the document layout is up‑to‑date before printing.
            report.UpdatePageLayout();

            // Print the populated document to the default printer.
            // The Print method is part of the Document class API.
            report.Print();
        }
    }
}
