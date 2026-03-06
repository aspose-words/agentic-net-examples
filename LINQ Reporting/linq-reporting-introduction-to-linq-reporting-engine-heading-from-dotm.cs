using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Simple data source class – any non‑dynamic, non‑anonymous type can be used.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTM template that already contains the heading
            // "<<[data.Title]>>" (or any other template syntax you prefer).
            Document template = new Document(@"Templates\LinqReportingTemplate.dotm");

            // Prepare the data source.
            var data = new ReportData
            {
                Title = "LINQ Reporting Introduction to LINQ Reporting Engine",
                Author = "Aspose.Words Team",
                Date = DateTime.Today
            };

            // Build the report using the LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            // The second overload allows us to reference the data source object itself via the name "data".
            engine.BuildReport(template, data, "data");

            // Save the populated document. The format is inferred from the extension (.docx).
            template.Save(@"Output\LinqReportingResult.docx");
        }
    }
}
