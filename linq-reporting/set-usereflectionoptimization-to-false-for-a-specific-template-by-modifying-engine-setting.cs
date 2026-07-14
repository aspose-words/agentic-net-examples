using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used in the template.
    public class Model
    {
        public string Name { get; set; } = "World";
    }

    // Wrapper that sets the static UseReflectionOptimization flag to false for the duration of the using block.
    public class EngineScope : IDisposable
    {
        public ReportingEngine Engine { get; }

        public EngineScope()
        {
            // Disable reflection optimization for this template.
            ReportingEngine.UseReflectionOptimization = false;
            Engine = new ReportingEngine();
        }

        public void Dispose()
        {
            // Restore the default value after the report is built.
            ReportingEngine.UseReflectionOptimization = true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document with a LINQ Reporting tag.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Hello <<[model.Name]>>!");

            // Prepare the data source.
            Model model = new Model { Name = "Aspose.Words" };

            // Build the report with reflection optimization disabled.
            using (var scope = new EngineScope())
            {
                scope.Engine.BuildReport(template, model, "model");
            }

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
