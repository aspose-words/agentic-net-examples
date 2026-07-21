using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "Default";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add a LINQ Reporting tag.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello <<[model.Name]>>!");

            // Prepare the data source.
            ReportModel model = new ReportModel { Name = "World" };

            // Disable reflection optimization for this specific report.
            ReportingEngine.UseReflectionOptimization = false;

            // Use ReportingEngine to build the report (no using statement because ReportingEngine is not IDisposable).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
